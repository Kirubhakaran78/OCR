#!/usr/bin/env python3
"""
Complete PDF Extractor with Sonnet 4 Vision
Extracts content from image-based PDFs and saves to output path
"""

import base64
import json
import fitz  # PyMuPDF
from pathlib import Path
from anthropic import Anthropic

class CompletePDFExtractor:
    """Complete PDF extractor using Sonnet 4 vision"""
    
    def __init__(self):
        # api_key = ""
        self.client = Anthropic(api_key=api_key)
        print("✅ Sonnet 4 client initialized with working model")
    
    def extract_pdf_with_vision(self, pdf_path, output_path="sonnet4_output"):
        """Extract PDF content using Sonnet 4 vision"""
        print(f"🔍 Starting PDF extraction: {pdf_path}")
        print(f"📁 Output directory: {output_path}")
        
        # Create output directory
        Path(output_path).mkdir(exist_ok=True)
        
        try:
            doc = fitz.open(pdf_path)
            results = {
                "pdf_path": pdf_path,
                "metadata": self.extract_metadata(pdf_path),
                "pages": [],
                "full_content": "",
                "summary": "",
                "structured_data": {}
            }
            
            print(f"📄 Processing {len(doc)} page(s)...")
            
            for page_num in range(len(doc)):
                print(f"   Processing page {page_num + 1}...")
                page = doc.load_page(page_num)
                
                # Convert page to high-quality image
                mat = fitz.Matrix(2.5, 2.5)  # High resolution
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("png")
                img_b64 = base64.b64encode(img_data).decode()
                
                # Analyze with Sonnet 4 vision
                page_content = self.analyze_page_vision(img_b64, page_num + 1)
                
                results["pages"].append({
                    "page_number": page_num + 1,
                    "content": page_content,
                    "image_size": len(img_data)
                })
                
                results["full_content"] += f"\n\n=== PAGE {page_num + 1} ===\n{page_content}"
            
            doc.close()
            
            # Generate summary and structured data
            results["summary"] = self.generate_summary(results["full_content"])
            results["structured_data"] = self.extract_structured_data(results["full_content"])
            
            # Save results to output path
            self.save_results(results, output_path)
            
            return results
            
        except Exception as e:
            print(f"❌ Error: {e}")
            return None
    
    def extract_metadata(self, pdf_path):
        """Extract PDF metadata"""
        try:
            doc = fitz.open(pdf_path)
            metadata = doc.metadata
            doc.close()
            return dict(metadata) if metadata else {}
        except:
            return {}
    
    def analyze_page_vision(self, img_b64, page_num):
        """Analyze page image with Sonnet 4 vision"""
        prompt = f"""
        Analyze this PDF page image and extract ALL visible content. This is page {page_num}.
        
        Please extract:
        1. ALL text content (headings, paragraphs, labels, captions, etc.)
        2. Document structure and sections
        3. Tables, lists, and structured data
        4. Numbers, dates, codes, references
        5. Any technical specifications or data
        6. Document title and main topics
        
        Be thorough and extract everything you can see. Format the response clearly with proper structure.
        """
        
        try:
            response = self.client.messages.create(
                model="claude-sonnet-4-20250514",  # Working model
                max_tokens=4000,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": img_b64}},
                        {"type": "text", "text": prompt}
                    ]
                }]
            )
            
            content = response.content[0].text
            print(f"   ✅ Extracted {len(content)} characters from page {page_num}")
            return content
            
        except Exception as e:
            print(f"   ❌ Vision analysis failed for page {page_num}: {e}")
            return f"Error analyzing page {page_num}: {str(e)}"
    
    def generate_summary(self, content):
        """Generate document summary"""
        if not content or len(content.strip()) < 50:
            return "No substantial content found for summary."
        
        try:
            prompt = f"""
            Based on this PDF content, provide a comprehensive summary:
            
            {content[:3000]}
            
            Include:
            1. Document title and type
            2. Main topics and sections
            3. Key information and data
            4. Document purpose and context
            """
            
            response = self.client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=1000,
                messages=[{"role": "user", "content": prompt}]
            )
            
            return response.content[0].text
            
        except Exception as e:
            return f"Summary generation failed: {str(e)}"
    
    def extract_structured_data(self, content):
        """Extract structured data from content"""
        try:
            prompt = f"""
            Extract structured data from this PDF content and return as JSON:
            
            {content[:2000]}
            
            Look for:
            - Document numbers, codes, references
            - Dates, timestamps, version numbers
            - Names, organizations, locations
            - Technical specifications
            - Key metrics or values
            - Table data
            
            Return valid JSON format.
            """
            
            response = self.client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=1500,
                messages=[{"role": "user", "content": prompt}]
            )
            
            try:
                return json.loads(response.content[0].text)
            except:
                return {"raw_extraction": response.content[0].text}
                
        except Exception as e:
            return {"error": f"Data extraction failed: {str(e)}"}
    
    def save_results(self, results, output_path):
        """Save all results to output directory"""
        base_name = Path(results["pdf_path"]).stem
        
        # Save complete JSON
        json_file = f"{output_path}/{base_name}_complete_extraction.json"
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"💾 Complete results: {json_file}")
        
        # Save readable text
        txt_file = f"{output_path}/{base_name}_extracted_content.txt"
        with open(txt_file, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("SONNET 4 PDF EXTRACTION RESULTS\n")
            f.write("=" * 80 + "\n\n")
            
            f.write(f"PDF File: {results['pdf_path']}\n")
            f.write(f"Pages Processed: {len(results['pages'])}\n")
            f.write(f"Total Content: {len(results['full_content'])} characters\n\n")
            
            f.write("=" * 80 + "\n")
            f.write("METADATA\n")
            f.write("=" * 80 + "\n")
            for key, value in results['metadata'].items():
                f.write(f"{key}: {value}\n")
            f.write("\n")
            
            f.write("=" * 80 + "\n")
            f.write("DOCUMENT SUMMARY\n")
            f.write("=" * 80 + "\n")
            f.write(results['summary'])
            f.write("\n\n")
            
            f.write("=" * 80 + "\n")
            f.write("STRUCTURED DATA\n")
            f.write("=" * 80 + "\n")
            f.write(json.dumps(results['structured_data'], indent=2))
            f.write("\n\n")
            
            f.write("=" * 80 + "\n")
            f.write("FULL EXTRACTED CONTENT\n")
            f.write("=" * 80 + "\n")
            f.write(results['full_content'])
        
        print(f"💾 Readable content: {txt_file}")
        
        # Save CSV format
        csv_file = f"{output_path}/{base_name}_structured_data.csv"
        self.save_as_csv(results, csv_file)
        print(f"💾 CSV data: {csv_file}")
        
        return {"json": json_file, "txt": txt_file, "csv": csv_file}
    
    def save_as_csv(self, results, csv_file):
        """Save structured data as CSV"""
        import csv
        
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(["category", "field", "value"])
            
            # Add metadata
            for key, value in results['metadata'].items():
                writer.writerow(["metadata", key, str(value)])
            
            # Add summary
            writer.writerow(["summary", "document_summary", results['summary']])
            
            # Add structured data
            if isinstance(results['structured_data'], dict):
                for key, value in results['structured_data'].items():
                    writer.writerow(["structured_data", key, str(value)])
            
            # Add page content
            for page in results['pages']:
                writer.writerow(["page_content", f"page_{page['page_number']}", page['content']])

def main():
    """Main execution"""
    pdf_path = "pg1.pdf"
    output_path = "sonnet4_output"
    
    print("🚀 Starting Complete PDF Extraction with Sonnet 4")
    print("=" * 70)
    
    extractor = CompletePDFExtractor()
    results = extractor.extract_pdf_with_vision(pdf_path, output_path)
    
    if results:
        print("\n" + "=" * 70)
        print("🎉 EXTRACTION COMPLETE!")
        print("=" * 70)
        print(f"📄 PDF: {pdf_path}")
        print(f"📊 Pages: {len(results['pages'])}")
        print(f"📝 Content: {len(results['full_content'])} characters")
        print(f"📁 Output: {output_path}/")
        print("\n📖 Content Preview:")
        print("-" * 50)
        preview = results['full_content'][:500] + "..." if len(results['full_content']) > 500 else results['full_content']
        print(preview)
    else:
        print("❌ Extraction failed")

if __name__ == "__main__":
    main()
