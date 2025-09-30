#!/usr/bin/env python3
"""
Enhanced PDF Extractor with Detailed Data Extraction
Extracts ALL content including tables, well data, and structured information
"""

import base64
import json
import csv
import fitz  # PyMuPDF
from pathlib import Path
from anthropic import Anthropic

class EnhancedPDFExtractor:
    """Enhanced PDF extractor with detailed table extraction"""
    
    def __init__(self):
        # api_key = ""
        self.client = Anthropic(api_key=api_key)
        print("✅ Sonnet 4 client initialized")
    
    def extract_pdf_with_vision(self, pdf_path, output_path="output"):
        """Extract PDF content with detailed analysis"""
        print(f"🔍 Starting Enhanced PDF extraction: {pdf_path}")
        print(f"📁 Output directory: {output_path}")
        
        Path(output_path).mkdir(exist_ok=True)
        
        try:
            doc = fitz.open(pdf_path)
            results = {
                "pdf_path": pdf_path,
                "metadata": self.extract_metadata(pdf_path),
                "pages": [],
                "well_plate_data": [],
                "standards_table": [],
                "settings": {},
                "samples_data": [],
                "full_content": "",
                "summary": ""
            }
            
            print(f"📄 Processing {len(doc)} page(s)...")
            
            for page_num in range(len(doc)):
                print(f"   Processing page {page_num + 1}...")
                page = doc.load_page(page_num)
                
                # Convert page to high-quality image
                mat = fitz.Matrix(3.0, 3.0)  # Higher resolution
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("png")
                img_b64 = base64.b64encode(img_data).decode()
                
                # Detailed analysis with Sonnet 4 vision
                page_analysis = self.analyze_page_detailed(img_b64, page_num + 1)
                
                results["pages"].append(page_analysis)
                results["full_content"] += f"\n\n=== PAGE {page_num + 1} ===\n{page_analysis['text_content']}"
            
            doc.close()
            
            # Extract structured components
            self.extract_well_plate_data(results)
            self.extract_standards_table(results)
            self.extract_settings(results)
            self.extract_samples_data(results)
            
            # Generate summary
            results["summary"] = self.generate_summary(results["full_content"])
            
            # Save in all formats
            self.save_all_formats(results, output_path)
            
            return results
            
        except Exception as e:
            print(f"❌ Error: {e}")
            import traceback
            traceback.print_exc()
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
    
    def analyze_page_detailed(self, img_b64, page_num):
        """Detailed page analysis with table extraction"""
        prompt = f"""
        Analyze this fluorescence assay PDF page (page {page_num}) and extract ALL data in detail.
        
        EXTRACT EVERYTHING YOU SEE:
        
        1. WELL PLATE DATA - Extract each well box with:
           - Well ID (A1, A2, B1, etc.)
           - Sample type (Std, Blank, Reference, Sample)
           - Concentration value
           - Raw value
           - Reduced value
           - Date and time
           
        2. SETTINGS TABLE - Extract all instrument settings:
           - Endpoint type
           - Wavelengths (Ex, Em, Cutoff)
           - PMT settings
           - Number of flashes
           - Any other parameters
           
        3. STANDARDS TABLE - Extract the complete table:
           - Sample names
           - Concentration values
           - Wells
           - Values
           - Back Calculated Concentration
           - Percent Back Calc
           
        4. READ INFORMATION:
           - Instrument model
           - ROM version
           - Start read time
           - Temperature
           - Operator name
           
        5. CALCULATED VALUES:
           - Any ratios or calculations shown
           
        6. DOCUMENT IDENTIFIERS:
           - Experiment number
           - QC information
           - Document status
           - Approval information
           
        Format your response as structured JSON with clear sections for each data type.
        Be extremely thorough - include every number, value, and label you can see.
        """
        
        try:
            response = self.client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=8000,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": img_b64}},
                        {"type": "text", "text": prompt}
                    ]
                }]
            )
            
            content = response.content[0].text
            
            # Try to parse as JSON, otherwise keep as text
            try:
                structured_data = json.loads(content)
            except:
                structured_data = {"raw_content": content}
            
            print(f"   ✅ Extracted {len(content)} characters from page {page_num}")
            
            return {
                "page_number": page_num,
                "text_content": content,
                "structured_data": structured_data
            }
            
        except Exception as e:
            print(f"   ❌ Analysis failed for page {page_num}: {e}")
            return {
                "page_number": page_num,
                "text_content": f"Error analyzing page {page_num}: {str(e)}",
                "structured_data": {}
            }
    
    def extract_well_plate_data(self, results):
        """Extract well plate data from analysis"""
        well_data = []
        for page in results['pages']:
            if 'well_plate_data' in page['structured_data']:
                well_data.extend(page['structured_data']['well_plate_data'])
            elif 'wells' in page['structured_data']:
                well_data.extend(page['structured_data']['wells'])
        results['well_plate_data'] = well_data
    
    def extract_standards_table(self, results):
        """Extract standards calibration table"""
        standards = []
        for page in results['pages']:
            if 'standards_table' in page['structured_data']:
                standards.extend(page['structured_data']['standards_table'])
            elif 'standards' in page['structured_data']:
                standards.extend(page['structured_data']['standards'])
        results['standards_table'] = standards
    
    def extract_settings(self, results):
        """Extract instrument settings"""
        for page in results['pages']:
            if 'settings' in page['structured_data']:
                results['settings'].update(page['structured_data']['settings'])
            elif 'instrument_settings' in page['structured_data']:
                results['settings'].update(page['structured_data']['instrument_settings'])
    
    def extract_samples_data(self, results):
        """Extract sample measurements"""
        samples = []
        for page in results['pages']:
            if 'samples' in page['structured_data']:
                samples.extend(page['structured_data']['samples'])
            elif 'sample_data' in page['structured_data']:
                samples.extend(page['structured_data']['sample_data'])
        results['samples_data'] = samples
    
    def generate_summary(self, content):
        """Generate document summary"""
        if not content or len(content.strip()) < 50:
            return "No substantial content found for summary."
        
        try:
            prompt = f"""
            Provide a concise summary of this fluorescence assay report:
            
            {content[:3000]}
            
            Include:
            1. Experiment type and purpose
            2. Key instrument settings
            3. Number of standards and samples
            4. Main findings or results
            """
            
            response = self.client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=1000,
                messages=[{"role": "user", "content": prompt}]
            )
            
            return response.content[0].text
            
        except Exception as e:
            return f"Summary generation failed: {str(e)}"
    
    def save_all_formats(self, results, output_path):
        """Save results in CSV, TXT, and JSON formats"""
        base_name = Path(results["pdf_path"]).stem
        
        # 1. Save JSON (complete data)
        json_file = f"{output_path}/{base_name}_complete.json"
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"💾 JSON saved: {json_file}")
        
        # 2. Save TXT (readable format)
        txt_file = f"{output_path}/{base_name}_content.txt"
        self.save_as_txt(results, txt_file)
        print(f"💾 TXT saved: {txt_file}")
        
        # 3. Save detailed CSV (structured data for Excel)
        csv_file = f"{output_path}/{base_name}_data.csv"
        self.save_as_detailed_csv(results, csv_file)
        print(f"💾 CSV saved: {csv_file}")
        
        print(f"\n📊 All formats saved in: {output_path}/")
        
        return {"json": json_file, "txt": txt_file, "csv": csv_file}
    
    def save_as_txt(self, results, txt_file):
        """Save as readable text file"""
        with open(txt_file, 'w', encoding='utf-8') as f:
            f.write("=" * 100 + "\n")
            f.write("FLUORESCENCE ASSAY - PDF EXTRACTION RESULTS\n")
            f.write("=" * 100 + "\n\n")
            
            f.write(f"PDF File: {results['pdf_path']}\n")
            f.write(f"Pages Processed: {len(results['pages'])}\n\n")
            
            # Metadata
            f.write("=" * 100 + "\n")
            f.write("PDF METADATA\n")
            f.write("=" * 100 + "\n")
            for key, value in results['metadata'].items():
                f.write(f"{key}: {value}\n")
            f.write("\n")
            
            # Settings
            if results['settings']:
                f.write("=" * 100 + "\n")
                f.write("INSTRUMENT SETTINGS\n")
                f.write("=" * 100 + "\n")
                f.write(json.dumps(results['settings'], indent=2))
                f.write("\n\n")
            
            # Well Plate Data
            if results['well_plate_data']:
                f.write("=" * 100 + "\n")
                f.write("WELL PLATE DATA\n")
                f.write("=" * 100 + "\n")
                for well in results['well_plate_data']:
                    f.write(json.dumps(well, indent=2))
                    f.write("\n")
                f.write("\n")
            
            # Standards Table
            if results['standards_table']:
                f.write("=" * 100 + "\n")
                f.write("STANDARDS CALIBRATION TABLE\n")
                f.write("=" * 100 + "\n")
                for std in results['standards_table']:
                    f.write(json.dumps(std, indent=2))
                    f.write("\n")
                f.write("\n")
            
            # Summary
            f.write("=" * 100 + "\n")
            f.write("DOCUMENT SUMMARY\n")
            f.write("=" * 100 + "\n")
            f.write(results['summary'])
            f.write("\n\n")
            
            # Full content
            f.write("=" * 100 + "\n")
            f.write("COMPLETE EXTRACTED CONTENT\n")
            f.write("=" * 100 + "\n")
            f.write(results['full_content'])
    
    def save_as_detailed_csv(self, results, csv_file):
        """Save as detailed CSV with all data sections"""
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            
            # Header
            writer.writerow(["Section", "Category", "Field", "Value", "Additional_Info"])
            writer.writerow([])
            
            # Metadata
            writer.writerow(["METADATA", "", "", "", ""])
            for key, value in results['metadata'].items():
                writer.writerow(["METADATA", "PDF Info", key, str(value), ""])
            writer.writerow([])
            
            # Settings
            writer.writerow(["INSTRUMENT SETTINGS", "", "", "", ""])
            for key, value in results['settings'].items():
                writer.writerow(["SETTINGS", "Instrument", key, str(value), ""])
            writer.writerow([])
            
            # Well Plate Data
            writer.writerow(["WELL PLATE DATA", "", "", "", ""])
            writer.writerow(["WELL DATA", "Well_ID", "Sample_Type", "Value", "Additional_Details"])
            for well in results['well_plate_data']:
                if isinstance(well, dict):
                    well_id = well.get('well', well.get('well_id', ''))
                    sample_type = well.get('type', well.get('sample_type', ''))
                    value = well.get('value', well.get('raw_value', ''))
                    details = json.dumps({k: v for k, v in well.items() if k not in ['well', 'type', 'value']})
                    writer.writerow(["WELL DATA", well_id, sample_type, str(value), details])
            writer.writerow([])
            
            # Standards Table
            writer.writerow(["STANDARDS CALIBRATION", "", "", "", ""])
            writer.writerow(["STANDARDS", "Sample", "Concentration", "Well", "Value", "Back_Calc"])
            for std in results['standards_table']:
                if isinstance(std, dict):
                    sample = std.get('sample', std.get('standard', ''))
                    conc = std.get('concentration', '')
                    well = std.get('well', '')
                    value = std.get('value', '')
                    back_calc = std.get('back_calc', std.get('percent_back_calc', ''))
                    writer.writerow(["STANDARDS", sample, str(conc), well, str(value), str(back_calc)])
            writer.writerow([])
            
            # Samples Data
            writer.writerow(["SAMPLES DATA", "", "", "", ""])
            for sample in results['samples_data']:
                if isinstance(sample, dict):
                    writer.writerow(["SAMPLE", 
                                   sample.get('name', ''), 
                                   sample.get('type', ''), 
                                   sample.get('value', ''),
                                   json.dumps(sample)])
            writer.writerow([])
            
            # Summary
            writer.writerow(["SUMMARY", "", "", "", ""])
            writer.writerow(["SUMMARY", "Document Summary", results['summary'], "", ""])
            writer.writerow([])
            
            # Page Content
            writer.writerow(["PAGE CONTENT", "", "", "", ""])
            for page in results['pages']:
                writer.writerow(["PAGE", f"Page_{page['page_number']}", 
                               page['text_content'][:200] + "...", "", ""])

def main():
    """Main execution"""
    pdf_path = "pg_data2.pdf"  # Change this to your PDF file
    output_path = "output"
    
    print("🚀 Starting Enhanced PDF Extraction")
    print("=" * 70)
    
    extractor = EnhancedPDFExtractor()
    results = extractor.extract_pdf_with_vision(pdf_path, output_path)
    
    if results:
        print("\n" + "=" * 70)
        print("🎉 EXTRACTION COMPLETE!")
        print("=" * 70)
        print(f"📄 PDF: {pdf_path}")
        print(f"📊 Pages: {len(results['pages'])}")
        print(f"🧪 Wells Extracted: {len(results['well_plate_data'])}")
        print(f"📈 Standards: {len(results['standards_table'])}")
        print(f"🔬 Samples: {len(results['samples_data'])}")
        print(f"\n💡 Next step: Run csv_to_excel_converter.py to create formatted Excel")
    else:
        print("❌ Extraction failed")

if __name__ == "__main__":
    main()