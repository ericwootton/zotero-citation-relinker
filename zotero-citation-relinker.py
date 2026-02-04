#!/usr/bin/env python3
"""
Zotero Citation Relinker

This script attempts to relink orphaned Zotero citations in a Word document
to items in your Zotero library using fuzzy matching.

Usage:
    python zotero_relinker.py document.docx [--zotero-path PATH] [--output OUTPUT.docx] [--threshold 80]

Requirements:
    pip install rapidfuzz python-docx --break-system-packages
"""

import argparse
import io
import json
import os
import re
import sqlite3
import sys
import zipfile

# Fix Windows console encoding to handle Unicode characters in citation data
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

try:
    from rapidfuzz import fuzz, process
except ImportError:
    print("Installing rapidfuzz...")
    os.system("pip install rapidfuzz --break-system-packages -q")
    from rapidfuzz import fuzz, process


# Word XML namespaces
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

# Register namespaces for ElementTree
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


@dataclass
class CitationItem:
    """Represents a single item within a Zotero citation"""
    uri: Optional[str] = None
    uris: list = field(default_factory=list)
    item_key: Optional[str] = None
    library_id: Optional[str] = None
    
    # Metadata for matching
    title: str = ""
    authors: list = field(default_factory=list)
    year: Optional[str] = None
    doi: Optional[str] = None
    isbn: Optional[str] = None
    
    # Raw data
    item_data: dict = field(default_factory=dict)
    
    # Orphan/match results
    is_orphaned: bool = False
    matched_library_item: Optional[dict] = None
    match_score: float = 0.0
    match_method: str = ""
    
    def author_string(self) -> str:
        """Get authors as a searchable string"""
        parts = []
        for author in self.authors:
            if isinstance(author, dict):
                if 'family' in author:
                    parts.append(author.get('family', ''))
                elif 'literal' in author:
                    parts.append(author.get('literal', ''))
            elif isinstance(author, str):
                parts.append(author)
        return " ".join(parts)
    
    def search_string(self) -> str:
        """Create a combined search string for fuzzy matching"""
        parts = [self.title, self.author_string()]
        if self.year:
            parts.append(str(self.year))
        return " ".join(filter(None, parts))


@dataclass
class Citation:
    """Represents a Zotero citation field in the document"""
    field_code: str
    citation_data: dict
    items: list  # List of CitationItem
    field_index: int  # Position in document
    is_orphaned: bool = False
    
    @classmethod
    def from_field_code(cls, field_code: str, field_index: int) -> Optional['Citation']:
        """Parse a Zotero field code into a Citation object"""
        # Extract JSON from field code
        match = re.search(r'ADDIN ZOTERO_ITEM CSL_CITATION\s*(\{.*\})', field_code, re.DOTALL)
        if not match:
            return None
        
        try:
            json_str = match.group(1)
            citation_data = json.loads(json_str)
        except json.JSONDecodeError as e:
            print(f"Warning: Failed to parse citation JSON: {e}")
            return None
        
        items = []
        for item_data in citation_data.get('citationItems', []):
            ci = CitationItem()
            ci.item_data = item_data
            
            # Extract URIs and key - check all URIs for a valid Zotero item key
            ci.uris = item_data.get('uris', [])
            ci.uri = ci.uris[0] if ci.uris else None
            for uri in ci.uris:
                key_match = re.search(r'/items/([A-Za-z0-9]+)$', uri)
                if key_match:
                    ci.item_key = key_match.group(1)
                    break
            
            # Extract metadata from itemData
            item_metadata = item_data.get('itemData', {})
            ci.title = item_metadata.get('title', '')
            ci.year = str(item_metadata.get('issued', {}).get('date-parts', [[None]])[0][0] or '')
            if not ci.year:
                ci.year = item_metadata.get('date', '')[:4] if item_metadata.get('date') else ''
            
            ci.authors = item_metadata.get('author', [])
            ci.doi = item_metadata.get('DOI', '')
            ci.isbn = item_metadata.get('ISBN', '')
            
            items.append(ci)
        
        return cls(
            field_code=field_code,
            citation_data=citation_data,
            items=items,
            field_index=field_index
        )


class ZoteroDatabase:
    """Interface to the local Zotero SQLite database"""
    
    def __init__(self, zotero_path: Optional[str] = None):
        self.db_path = self._find_database(zotero_path)
        self.items = []
        self.items_by_key = {}
        self.items_by_doi = {}
        self.items_by_isbn = {}
        self.user_id = None
        self.local_user_key = None

        if self.db_path:
            self._load_items()
    
    def _find_database(self, custom_path: Optional[str] = None) -> Optional[Path]:
        """Find the Zotero database file"""
        if custom_path:
            path = Path(custom_path)
            if path.is_file():
                return path
            elif path.is_dir():
                db_path = path / 'zotero.sqlite'
                if db_path.exists():
                    return db_path
        
        # Try common locations
        home = Path.home()
        possible_paths = [
            home / 'Zotero' / 'zotero.sqlite',
            home / '.zotero' / 'zotero' / 'zotero.sqlite',
            home / 'snap' / 'zotero-snap' / 'common' / 'Zotero' / 'zotero.sqlite',
            home / 'Library' / 'Application Support' / 'Zotero' / 'zotero.sqlite',  # macOS
            Path(os.environ.get('APPDATA', '')) / 'Zotero' / 'Zotero' / 'zotero.sqlite',  # Windows
        ]
        
        for path in possible_paths:
            if path.exists():
                return path
        
        return None
    
    def _load_items(self):
        """Load items from the Zotero database"""
        if not self.db_path:
            return
        
        # Make a copy of the database to avoid locking issues
        import shutil
        import tempfile
        
        temp_db = Path(tempfile.gettempdir()) / 'zotero_temp.sqlite'
        shutil.copy2(self.db_path, temp_db)
        
        try:
            conn = sqlite3.connect(temp_db)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()

            # Get user ID and local user key for URI construction
            cursor.execute("SELECT key, value FROM settings WHERE setting='account'")
            for row in cursor.fetchall():
                if row['key'] == 'userID':
                    self.user_id = row['value']
                elif row['key'] == 'localUserKey':
                    self.local_user_key = row['value']

            # Get all items with their metadata
            query = """
            SELECT 
                i.itemID,
                i.key,
                i.libraryID,
                it.typeName as itemType,
                (SELECT value FROM itemDataValues idv 
                 JOIN itemData id ON idv.valueID = id.valueID 
                 JOIN fields f ON id.fieldID = f.fieldID 
                 WHERE id.itemID = i.itemID AND f.fieldName = 'title') as title,
                (SELECT value FROM itemDataValues idv 
                 JOIN itemData id ON idv.valueID = id.valueID 
                 JOIN fields f ON id.fieldID = f.fieldID 
                 WHERE id.itemID = i.itemID AND f.fieldName = 'date') as date,
                (SELECT value FROM itemDataValues idv 
                 JOIN itemData id ON idv.valueID = id.valueID 
                 JOIN fields f ON id.fieldID = f.fieldID 
                 WHERE id.itemID = i.itemID AND f.fieldName = 'DOI') as doi,
                (SELECT value FROM itemDataValues idv 
                 JOIN itemData id ON idv.valueID = id.valueID 
                 JOIN fields f ON id.fieldID = f.fieldID 
                 WHERE id.itemID = i.itemID AND f.fieldName = 'ISBN') as isbn
            FROM items i
            JOIN itemTypes it ON i.itemTypeID = it.itemTypeID
            WHERE i.itemID NOT IN (SELECT itemID FROM deletedItems)
            AND it.typeName != 'attachment'
            AND it.typeName != 'note'
            """
            
            cursor.execute(query)
            rows = cursor.fetchall()
            
            for row in rows:
                item = dict(row)
                
                # Get authors/creators
                cursor.execute("""
                    SELECT c.firstName, c.lastName, ct.creatorType
                    FROM itemCreators ic
                    JOIN creators c ON ic.creatorID = c.creatorID
                    JOIN creatorTypes ct ON ic.creatorTypeID = ct.creatorTypeID
                    WHERE ic.itemID = ?
                    ORDER BY ic.orderIndex
                """, (item['itemID'],))
                
                creators = cursor.fetchall()
                item['authors'] = [{'given': c['firstName'], 'family': c['lastName']} 
                                   for c in creators if c['creatorType'] in ('author', 'editor')]
                
                # Extract year from date
                if item.get('date'):
                    year_match = re.search(r'\d{4}', item['date'])
                    item['year'] = year_match.group(0) if year_match else None
                else:
                    item['year'] = None
                
                # Create search string
                author_str = " ".join(a.get('family', '') for a in item['authors'])
                item['search_string'] = f"{item.get('title', '')} {author_str} {item.get('year', '')}"
                
                self.items.append(item)
                self.items_by_key[item['key']] = item
                
                if item.get('doi'):
                    self.items_by_doi[item['doi'].lower().strip()] = item
                if item.get('isbn'):
                    # Normalize ISBN
                    isbn_clean = re.sub(r'[^0-9X]', '', item['isbn'].upper())
                    self.items_by_isbn[isbn_clean] = item
            
            conn.close()
        finally:
            if temp_db.exists():
                temp_db.unlink()
        
        print(f"Loaded {len(self.items)} items from Zotero library")
    
    def find_match(self, citation_item: CitationItem, threshold: int = 80) -> Optional[dict]:
        """Find a matching library item for a citation item"""
        # Try exact DOI match first
        if citation_item.doi:
            doi_clean = citation_item.doi.lower().strip()
            if doi_clean in self.items_by_doi:
                citation_item.match_method = "DOI"
                citation_item.match_score = 100
                return self.items_by_doi[doi_clean]
        
        # Try exact ISBN match
        if citation_item.isbn:
            isbn_clean = re.sub(r'[^0-9X]', '', citation_item.isbn.upper())
            if isbn_clean in self.items_by_isbn:
                citation_item.match_method = "ISBN"
                citation_item.match_score = 100
                return self.items_by_isbn[isbn_clean]
        
        # Try fuzzy matching on title + author + year
        search_str = citation_item.search_string()
        if not search_str.strip():
            return None
        
        choices = [(item['search_string'], item['key']) for item in self.items]
        
        if not choices:
            return None
        
        # Use token_set_ratio for better matching with reordered words
        result = process.extractOne(
            search_str,
            [c[0] for c in choices],
            scorer=fuzz.token_set_ratio
        )
        
        if result and result[1] >= threshold:
            match_key = choices[result[2]][1]
            citation_item.match_method = "fuzzy"
            citation_item.match_score = result[1]
            return self.items_by_key[match_key]
        
        # Try matching on title alone with higher threshold
        if citation_item.title:
            title_choices = [(item.get('title', ''), item['key']) for item in self.items if item.get('title')]
            result = process.extractOne(
                citation_item.title,
                [c[0] for c in title_choices],
                scorer=fuzz.token_set_ratio
            )
            
            if result and result[1] >= 90:  # Higher threshold for title-only match
                match_key = title_choices[result[2]][1]
                citation_item.match_method = "title_only"
                citation_item.match_score = result[1]
                return self.items_by_key[match_key]
        
        return None


def extract_citations_from_docx(docx_path: str) -> list[Citation]:
    """Extract all Zotero citations from a Word document"""
    citations = []
    
    with zipfile.ZipFile(docx_path, 'r') as zf:
        # Read document.xml
        with zf.open('word/document.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
        
        # Find all field codes (w:instrText elements)
        field_index = 0
        
        # Zotero uses complex fields: w:fldChar (begin) -> w:instrText -> w:fldChar (end)
        # We need to collect all instrText content between begin and end markers
        
        current_field = []
        in_field = False
        
        for elem in root.iter():
            # Check for field character
            if elem.tag == f"{{{NAMESPACES['w']}}}fldChar":
                fld_type = elem.get(f"{{{NAMESPACES['w']}}}fldCharType")
                if fld_type == 'begin':
                    in_field = True
                    current_field = []
                elif fld_type == 'end' and in_field:
                    in_field = False
                    field_code = ''.join(current_field)
                    
                    if 'ADDIN ZOTERO_ITEM' in field_code:
                        citation = Citation.from_field_code(field_code, field_index)
                        if citation:
                            citations.append(citation)
                            field_index += 1
                    
                    current_field = []
            
            # Collect instruction text
            elif in_field and elem.tag == f"{{{NAMESPACES['w']}}}instrText":
                if elem.text:
                    current_field.append(elem.text)
    
    return citations


def check_orphaned_status(citations: list[Citation], db: ZoteroDatabase):
    """Check which citations are orphaned (not linked to library)"""
    for citation in citations:
        for item in citation.items:
            # A citation is orphaned if it has no key or the key doesn't exist in the library
            if item.item_key and item.item_key in db.items_by_key:
                item.is_orphaned = False
            else:
                item.is_orphaned = True

        # Citation is orphaned if any of its items are orphaned
        citation.is_orphaned = any(item.is_orphaned for item in citation.items)


def generate_report(citations: list[Citation], threshold: int = 80) -> str:
    """Generate a report of orphaned citations and potential matches"""
    lines = []
    lines.append("=" * 80)
    lines.append("ZOTERO CITATION RELINKER REPORT")
    lines.append("=" * 80)
    lines.append("")
    
    total_citations = len(citations)
    total_items = sum(len(c.items) for c in citations)
    orphaned_citations = [c for c in citations if c.is_orphaned]
    orphaned_items = sum(len([i for i in c.items if i.is_orphaned]) for c in orphaned_citations)
    
    lines.append(f"Total citations in document: {total_citations}")
    lines.append(f"Total citation items: {total_items}")
    lines.append(f"Orphaned citations: {len(orphaned_citations)}")
    lines.append(f"Orphaned items: {orphaned_items}")
    lines.append("")
    
    if not orphaned_citations:
        lines.append("[OK] No orphaned citations found! All citations are linked to your library.")
        return "\n".join(lines)
    
    lines.append("-" * 80)
    lines.append("ORPHANED CITATIONS AND POTENTIAL MATCHES")
    lines.append("-" * 80)
    lines.append("")
    
    matched_count = 0
    unmatched_count = 0
    
    for i, citation in enumerate(orphaned_citations, 1):
        for j, item in enumerate(citation.items):
            if not item.is_orphaned:
                continue
            
            lines.append(f"[{i}.{j+1}] ORPHANED CITATION:")
            lines.append(f"    Title:   {item.title or '(no title)'}")
            lines.append(f"    Authors: {item.author_string() or '(no authors)'}")
            lines.append(f"    Year:    {item.year or '(no year)'}")
            if item.doi:
                lines.append(f"    DOI:     {item.doi}")
            lines.append("")
            
            match = item.matched_library_item

            if match:
                matched_count += 1

                lines.append(f"    [MATCH] POTENTIAL MATCH FOUND ({item.match_method}, score: {item.match_score}%):")
                lines.append(f"      Library Key: {match['key']}")
                lines.append(f"      Title:       {match.get('title', '(no title)')}")
                author_str = " ".join(a.get('family', '') for a in match.get('authors', []))
                lines.append(f"      Authors:     {author_str or '(no authors)'}")
                lines.append(f"      Year:        {match.get('year', '(no year)')}")
            else:
                unmatched_count += 1
                lines.append(f"    [NONE] NO MATCH FOUND (threshold: {threshold}%)")
            
            lines.append("")
    
    lines.append("-" * 80)
    lines.append("SUMMARY")
    lines.append("-" * 80)
    lines.append(f"Matches found:     {matched_count}")
    lines.append(f"No matches found:  {unmatched_count}")
    
    if matched_count > 0:
        lines.append("")
        lines.append("To relink these citations, you can:")
        lines.append("1. Use the --output flag to generate a new document with updated links")
        lines.append("2. Manually update citations in Word using the Zotero plugin")
    
    return "\n".join(lines)


def update_docx_citations(docx_path: str, output_path: str, citations: list[Citation], db: ZoteroDatabase):
    """Create a new DOCX with updated citation links"""
    import shutil
    import tempfile

    # Build a mapping of old URIs to new URIs from matched citations
    uri_replacements = {}
    for citation in citations:
        if not citation.is_orphaned:
            continue
        for item in citation.items:
            if not item.is_orphaned or not item.matched_library_item:
                continue
            lib_item = item.matched_library_item
            new_key = lib_item['key']
            # Construct proper URI with the user's actual Zotero user ID
            if db.user_id:
                new_uri = f"http://zotero.org/users/{db.user_id}/items/{new_key}"
            elif db.local_user_key:
                new_uri = f"http://zotero.org/users/local/{db.local_user_key}/items/{new_key}"
            else:
                new_uri = f"http://zotero.org/users/local/items/{new_key}"

            # Map each old URI for this item to the new URI
            for old_uri in item.uris:
                uri_replacements[old_uri] = new_uri

    if not uri_replacements:
        print("No matched citations to update.")
        return

    # Copy to temp location
    temp_dir = tempfile.mkdtemp()
    temp_docx = Path(temp_dir) / 'temp.docx'
    shutil.copy2(docx_path, temp_docx)

    # Extract the docx
    extract_dir = Path(temp_dir) / 'extracted'
    with zipfile.ZipFile(temp_docx, 'r') as zf:
        zf.extractall(extract_dir)

    # Read and modify document.xml
    doc_xml_path = extract_dir / 'word' / 'document.xml'
    with open(doc_xml_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Replace each old URI with the new one directly in the raw XML text
    # This preserves the original JSON formatting (no re-serialization needed)
    replacements_made = 0
    for old_uri, new_uri in uri_replacements.items():
        if old_uri in content:
            content = content.replace(old_uri, new_uri)
            replacements_made += 1

    print(f"Replaced {replacements_made} URI(s) in document")

    # Write updated document.xml
    with open(doc_xml_path, 'w', encoding='utf-8') as f:
        f.write(content)

    # Repackage the docx
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                file_path = Path(root) / file
                arcname = file_path.relative_to(extract_dir)
                zf.write(file_path, arcname)

    # Cleanup
    shutil.rmtree(temp_dir)

    print(f"Updated document saved to: {output_path}")


def generate_manual_relink_script(citations: list[Citation], output_path: str):
    """Generate a helper document with instructions for manual relinking"""
    lines = []
    lines.append("MANUAL RELINKING GUIDE")
    lines.append("=" * 60)
    lines.append("")
    lines.append("For each orphaned citation below, follow these steps in Word:")
    lines.append("1. Click on the citation in your document")
    lines.append("2. Click 'Add/Edit Citation' in the Zotero toolbar")
    lines.append("3. Delete the orphaned item (click X on the bubble)")
    lines.append("4. Search for and add the matching item from your library")
    lines.append("")
    lines.append("-" * 60)
    lines.append("")
    
    for citation in citations:
        if not citation.is_orphaned:
            continue
        
        for item in citation.items:
            if not item.is_orphaned:
                continue
            
            lines.append(f"ORPHANED: {item.title or '(no title)'}")
            lines.append(f"  Authors: {item.author_string()}")
            lines.append(f"  Year: {item.year}")
            
            if item.matched_library_item:
                match = item.matched_library_item
                lines.append(f"  -> SEARCH FOR: \"{match.get('title', '')}\"")
                lines.append(f"    Library Key: {match['key']}")
            else:
                lines.append(f"  -> NO AUTOMATIC MATCH - Search manually")
            
            lines.append("")
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))
    
    print(f"Manual relinking guide saved to: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description='Relink orphaned Zotero citations in Word documents',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Analyze a document and show matches
  python zotero_relinker.py thesis.docx

  # Specify Zotero data directory
  python zotero_relinker.py thesis.docx --zotero-path ~/Zotero

  # Generate updated document
  python zotero_relinker.py thesis.docx --output thesis_fixed.docx

  # Lower match threshold for more matches
  python zotero_relinker.py thesis.docx --threshold 70
        """
    )
    
    parser.add_argument('docx_file', help='Path to the Word document')
    parser.add_argument('--zotero-path', '-z', help='Path to Zotero data directory or zotero.sqlite')
    parser.add_argument('--output', '-o', help='Output path for updated document')
    parser.add_argument('--threshold', '-t', type=int, default=80, 
                        help='Fuzzy match threshold (0-100, default: 80)')
    parser.add_argument('--guide', '-g', help='Generate manual relinking guide to this file')
    
    args = parser.parse_args()
    
    # Check input file exists
    if not os.path.exists(args.docx_file):
        print(f"Error: File not found: {args.docx_file}")
        sys.exit(1)
    
    # Initialize Zotero database connection
    print("Connecting to Zotero database...")
    db = ZoteroDatabase(args.zotero_path)
    
    if not db.db_path:
        print("Warning: Could not find Zotero database.")
        print("Please specify the path with --zotero-path")
        print("Common locations:")
        print("  - ~/Zotero/zotero.sqlite (Linux/Mac)")
        print("  - %APPDATA%/Zotero/Zotero/zotero.sqlite (Windows)")
        sys.exit(1)
    
    print(f"Using database: {db.db_path}")
    
    # Extract citations from document
    print(f"\nAnalyzing document: {args.docx_file}")
    citations = extract_citations_from_docx(args.docx_file)
    
    if not citations:
        print("No Zotero citations found in the document.")
        sys.exit(0)
    
    print(f"Found {len(citations)} Zotero citations")
    
    # Check orphaned status
    check_orphaned_status(citations, db)
    
    # Find matches for orphaned citations
    for citation in citations:
        for item in citation.items:
            if item.is_orphaned:
                match = db.find_match(item, args.threshold)
                if match:
                    item.matched_library_item = match

    # Generate and print report
    report = generate_report(citations, args.threshold)
    print("\n" + report)
    
    # Save report
    report_path = Path(args.docx_file).stem + "_relink_report.txt"
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(report)
    print(f"\nReport saved to: {report_path}")
    
    # Generate manual guide if requested
    if args.guide:
        generate_manual_relink_script(citations, args.guide)
    
    # Generate updated document if requested
    if args.output:
        print("\nGenerating updated document...")
        update_docx_citations(args.docx_file, args.output, citations, db)
        print("\n** IMPORTANT: The updated document may need manual verification.")
        print("   Open it in Word with Zotero running and click 'Refresh' to verify links.")


if __name__ == '__main__':
    main()
