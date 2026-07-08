#!/usr/bin/env python3
"""
Automated system for extracting and formatting cosmetic ingredient data from Word documents.
Processes a messy base document with multiple ingredients and creates standardized 
single-ingredient documents with consistent table structure.
"""

import os
import re
from typing import List, Dict, Tuple, Optional
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pathlib import Path


class IngredientExtractor:
    """Extracts ingredient data from a Word document."""
    
    def __init__(self, base_doc_path: str, template_doc_path: str):
        """
        Initialize the extractor with base and template documents.
        
        Args:
            base_doc_path: Path to the messy base document
            template_doc_path: Path to the template document with standardized format
        """
        self.base_doc = Document(base_doc_path)
        self.template_doc = Document(template_doc_path)
        self.ingredients = []
        self.output_dir = Path("output_ingredients")
        self.output_dir.mkdir(exist_ok=True)
    
    def extract_ingredient_sections(self) -> List[Dict]:
        """
        Extract ingredient sections from the base document.
        Identifies ingredient names and their associated content/tables.
        
        Returns:
            List of dictionaries containing ingredient data
        """
        ingredients = []
        current_ingredient = None
        current_content = []
        
        for para in self.base_doc.paragraphs:
            text = para.text.strip()
            
            # Detect ingredient name (usually uppercase or specific pattern)
            if self._is_ingredient_header(text):
                # Save previous ingredient if exists
                if current_ingredient:
                    ingredients.append({
                        'name': current_ingredient,
                        'content': current_content,
                        'paragraphs': len(current_content)
                    })
                
                current_ingredient = text
                current_content = []
            elif current_ingredient and text:
                current_content.append(text)
        
        # Don't forget the last ingredient
        if current_ingredient:
            ingredients.append({
                'name': current_ingredient,
                'content': current_content,
                'paragraphs': len(current_content)
            })
        
        self.ingredients = ingredients
        return ingredients
    
    def extract_tables_by_ingredient(self) -> Dict[str, List]:
        """
        Extract tables from the base document organized by ingredient.
        
        Returns:
            Dictionary mapping ingredient names to their tables
        """
        ingredient_tables = {ing['name']: [] for ing in self.ingredients}
        current_ingredient = None
        
        for para in self.base_doc.paragraphs:
            text = para.text.strip()
            if self._is_ingredient_header(text):
                current_ingredient = text
        
        # Extract tables
        for table in self.base_doc.tables:
            if current_ingredient and current_ingredient in ingredient_tables:
                table_data = self._extract_table_data(table)
                ingredient_tables[current_ingredient].append(table_data)
        
        return ingredient_tables
    
    def _extract_table_data(self, table) -> List[List[str]]:
        """
        Convert a table to a list of lists containing cell data.
        
        Args:
            table: A docx table object
            
        Returns:
            List of rows, where each row is a list of cell values
        """
        table_data = []
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            table_data.append(row_data)
        return table_data
    
    def _is_ingredient_header(self, text: str) -> bool:
        """
        Detect if text is an ingredient name header.
        Heuristics: uppercase, no lowercase, 2-10 words, etc.
        
        Args:
            text: Text to check
            
        Returns:
            True if text appears to be an ingredient header
        """
        if not text or len(text) < 3:
            return False
        
        # Simple heuristics
        word_count = len(text.split())
        is_uppercase = text.isupper() or text.istitle()
        has_min_length = len(text) >= 5
        not_sentence = not any(text.endswith(c) for c in ['.', '!', '?'])
        
        return is_uppercase and has_min_length and word_count <= 10 and not_sentence
    
    def create_standardized_document(self, ingredient_name: str, 
                                     ingredient_data: Dict) -> Document:
        """
        Create a standardized document for a single ingredient using the template.
        
        Args:
            ingredient_name: Name of the ingredient
            ingredient_data: Dictionary with ingredient information
            
        Returns:
            A new Document object with standardized format
        """
        # Clone the template document
        new_doc = Document(self.template_doc)
        
        # Update document with ingredient-specific data
        # This assumes the template has placeholders or specific structure
        self._populate_document(new_doc, ingredient_name, ingredient_data)
        
        return new_doc
    
    def _populate_document(self, doc: Document, ingredient_name: str, 
                          ingredient_data: Dict) -> None:
        """
        Populate a document with ingredient-specific data.
        
        Args:
            doc: Document to populate
            ingredient_name: Name of the ingredient
            ingredient_data: Ingredient data to insert
        """
        # Add/update title
        if doc.paragraphs:
            doc.paragraphs[0].text = ingredient_name
        
        # Add ingredient data as content
        for content_item in ingredient_data.get('content', []):
            doc.add_paragraph(content_item)
        
        # Add tables if available
        for table_data in ingredient_data.get('tables', []):
            self._add_formatted_table(doc, table_data)
    
    def _add_formatted_table(self, doc: Document, table_data: List[List[str]]) -> None:
        """
        Add a formatted table to the document.
        
        Args:
            doc: Document to add table to
            table_data: Table data as list of lists
        """
        if not table_data:
            return
        
        rows = len(table_data)
        cols = len(table_data[0]) if table_data else 0
        
        if rows == 0 or cols == 0:
            return
        
        table = doc.add_table(rows=rows, cols=cols)
        table.style = 'Light Grid Accent 1'
        
        for i, row_data in enumerate(table_data):
            for j, cell_value in enumerate(row_data):
                cell = table.rows[i].cells[j]
                cell.text = str(cell_value)
    
    def process_all_ingredients(self) -> None:
        """
        Process all ingredients and create standardized documents.
        Saves each ingredient to a separate file.
        """
        print(f"Processing {len(self.ingredients)} ingredients...")
        
        for ingredient in self.ingredients:
            ingredient_name = ingredient['name']
            print(f"  Processing: {ingredient_name}")
            
            # Prepare ingredient data
            ingredient_data = {
                'name': ingredient_name,
                'content': ingredient['content'],
                'tables': ingredient.get('tables', [])
            }
            
            # Create standardized document
            new_doc = self.create_standardized_document(ingredient_name, ingredient_data)
            
            # Save document
            output_filename = self._sanitize_filename(ingredient_name)
            output_path = self.output_dir / f"{output_filename}.docx"
            
            try:
                new_doc.save(str(output_path))
                print(f"    ✓ Saved: {output_path}")
            except Exception as e:
                print(f"    ✗ Error saving {output_filename}: {str(e)}")
    
    def _sanitize_filename(self, filename: str) -> str:
        """
        Sanitize ingredient name for use as filename.
        
        Args:
            filename: Original filename/ingredient name
            
        Returns:
            Sanitized filename safe for file system
        """
        # Remove invalid characters
        invalid_chars = r'[<>:"/\\|?*]'
        sanitized = re.sub(invalid_chars, '', filename)
        # Replace spaces with underscores
        sanitized = sanitized.replace(' ', '_')
        # Limit length
        sanitized = sanitized[:255]
        return sanitized
    
    def print_summary(self) -> None:
        """Print a summary of extracted ingredients."""
        print("\n" + "="*60)
        print("EXTRACTION SUMMARY")
        print("="*60)
        print(f"Total ingredients found: {len(self.ingredients)}\n")
        
        for i, ingredient in enumerate(self.ingredients, 1):
            print(f"{i}. {ingredient['name']}")
            print(f"   - Content paragraphs: {ingredient['paragraphs']}")
            print(f"   - Tables: {len(ingredient.get('tables', []))}")
        
        print("="*60 + "\n")


def main():
    """Main execution function."""
    
    # Configuration
    BASE_DOC = "BAZA.docx"
    TEMPLATE_DOC = "00_WZÓR.docx"
    
    # Check if files exist
    if not Path(BASE_DOC).exists():
        print(f"Error: {BASE_DOC} not found")
        return
    
    if not Path(TEMPLATE_DOC).exists():
        print(f"Error: {TEMPLATE_DOC} not found")
        return
    
    # Initialize extractor
    print("Initializing ingredient extractor...")
    extractor = IngredientExtractor(BASE_DOC, TEMPLATE_DOC)
    
    # Extract ingredients
    print("Extracting ingredients from base document...")
    ingredients = extractor.extract_ingredient_sections()
    
    # Extract tables
    print("Extracting tables...")
    tables_by_ingredient = extractor.extract_tables_by_ingredient()
    
    # Attach tables to ingredients
    for ingredient in extractor.ingredients:
        ingredient['tables'] = tables_by_ingredient.get(ingredient['name'], [])
    
    # Print summary
    extractor.print_summary()
    
    # Process all ingredients
    print("Creating standardized documents...")
    extractor.process_all_ingredients()
    
    print(f"✓ All documents saved to: {extractor.output_dir}/")


if __name__ == "__main__":
    main()
