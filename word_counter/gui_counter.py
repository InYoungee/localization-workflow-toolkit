import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import json
from docx import Document
import PyPDF2
from datetime import datetime
import pandas as pd


en_ko_rate = 0.15
class WordCounterGUI:
	def __init__(self, root):
		self.root = root
		self.root.title("Localization Word Counter")
		self.root.geometry("750x700")

		self.results = []
		self.create_widgets()

	def create_widgets(self):
		title = tk.Label(self.root, text="Localization Word Counter", font=("Arial", 18, "bold"))
		title.pack(pady=20)

		instructions = tk.Label(self.root, text="Drag and drop files here, or click 'Browse Files'",
									font=("Arial", 12))
		instructions.pack(pady=10)

		self.drop_frame = tk.Frame(self.root, width=600, height="200", bg="#e0e0e0", relief=tk.SUNKEN, bd=2)
		self.drop_frame.pack(pady=20)
		self.drop_frame.pack_propagate(False)

		self.drop_label = tk.Label(self.drop_frame, text="Drop files here\n\nSupported: JSON, XML, XLF, DOCX, PDF",
									font=("Arial", 11), bg="#e0e0e0", fg="#666")
		self.drop_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

		try:
			self.drop_frame.drop_target_register(DND_FILES)
			self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
		except:
			pass

		browse_btn = tk.Button(self.root, text="Browse Files", command=self.browse_files, font=("Arial", 12),
								   bg="#4CAF50", fg="black",padx=20, pady=10)
		browse_btn.pack(pady=10)

		result_frame = tk.Frame(self.root)
		result_frame.pack(pady=10, fill=tk.BOTH, expand=True)

		scrollbar = tk.Scrollbar(result_frame)
		scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

		self.result_text = tk.Text(result_frame, height=15, width=80,yscrollcommand=scrollbar.set, font=("Courier", 10))
		self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
		scrollbar.config(command=self.result_text.yview)

		export_btn = tk.Button(self.root, text="Export to Excel", command=self.export_to_excel, font=("Arial", 12), bg="#2196F3", fg="black", padx=20, pady=10, cursor="hand2")
		export_btn.pack(pady=10)

	def on_drop(self, event):
		files = self.root.tk.splitlist(event.data)
		self.process_files(files)

	def browse_files(self):
		files = filedialog.askopenfilenames(
			title="Select files",
			filetypes=[
				("All Supported", "*.json *.xml *.xlf *.docx *.pdf"),
				("JSON files", "*.json"),
				("XML files", "*.xml"),
				("XLF files", "*.xlf"),
				("Word files", "*.docx"),
				("PDF files", "*.pdf"),
				("All files", "*.*")
			]
		)
		if files:
			self.process_files(files)

	def process_files(self, files):
		self.results = []
		self.result_text.delete(1.0, tk.END)
		self.result_text.insert(tk.END, "Processing files...\n\n")
		self.root.update()

		for filepath in files:
			result = self.count_words_in_file(filepath)
			if result:
				self.results.append(result)
		self.display_results()

	def count_words_in_file(self, filepath):
		filename = os.path.basename(filepath)
		_, ext = os.path.splitext(filename)
		ext = ext.lower()

		try:
			if ext == ".json":
				text = self.extract_from_json(filepath)
				file_type = 'JSON'
			elif ext in ['.xml', '.xlf']:
				text = self.extract_from_xml(filepath)
				file_type = 'XML/XLF'
			elif ext == '.docx':
				text = self.extract_from_docx(filepath)
				text = self.extract_from_docx(filepath)
				file_type = 'DOCX'
			elif ext =='.pdf':
				text = self.extract_from_pdf(filepath)
				file_type = 'PDF'
			else:
				self.result_text.insert(tk.END, f"⚠ Unsupported: {filename}\n")
				return None

			words = len(text.split()) if text else 0
			return {
				'filename': filename,
				'file_type': file_type,
				'words': words
			}

		except Exception as e:
			self.result_text.insert(tk.END, f"X Error with {filename}: {str(e)}\n")
			return None

	def extract_from_json(self, filepath):
		def get_strings(obj):
			text = []
			if isinstance(obj, dict):
				for value in obj.values():
					text.extend(get_strings(value))
			elif isinstance(obj, list):
				for item in obj:
					text.extend(get_strings(item))
			elif isinstance(obj, str):
				text.append(obj)
			return text

		with open(filepath, 'r', encoding='utf-8') as f:
			data = json.load(f)
		return ' '.join(get_strings(data))

	def extract_from_pdf(self, filepath):
		text = []
		with open(filepath, 'rb') as f:
			pdf = PyPDF2.PdfReader(f)
			for page in pdf.pages:
				text.append(page.extract_text())
		return ' '.join(text)

	def display_results(self):
		self.result_text.delete(1.0, tk.END)

		if not self.results:
			self.result_text.insert(tk.END, "No files processed successfully.\n")
			return

		# Header
		self.result_text.insert(tk.END, "=" * 70 + "\n")
		self.result_text.insert(tk.END, "WORD COUNT RESULTS\n")
		self.result_text.insert(tk.END, "=" * 70 + "\n\n")

		# Table header
		self.result_text.insert(tk.END, f"{'File Name':<35} {'Type':<10} {'Words':>10}\n")
		self.result_text.insert(tk.END, "-" * 70 + "\n")

		# Results
		total_words = 0
		for result in self.results:
			line = f"{result['filename']:<35} {result['file_type']:<10} {result['words']:>10,}\n"
			self.result_text.insert(tk.END, line)
			total_words += result['words']

		# Total
		self.result_text.insert(tk.END, "-" * 70 + "\n")
		self.result_text.insert(tk.END, f"{'TOTAL':<35} {'':<10} {total_words:>10,}\n")
		self.result_text.insert(tk.END, "=" * 70 + "\n\n")

		# Cost estimate
		cost_per_word = en_ko_rate
		estimated_cost = total_words * cost_per_word
		estimated_hours = total_words / 250

		self.result_text.insert(tk.END, f"Files processed: {len(self.results)}\n")
		self.result_text.insert(tk.END, f"Total words: {total_words:,}\n")
		self.result_text.insert(tk.END, f"Estimated cost: ${estimated_cost:,.2f} (at $0.15/word)\n")
		self.result_text.insert(tk.END, f"Estimated time: {estimated_hours:.1f} hours (at 250 words/hour)\n")

	def export_to_excel(self):
		"""Export results to Excel file"""
		if not self.results:
			messagebox.showwarning("No Data", "No results to export. Please process files first.")
			return

		try:
			# Check if pandas is available
			try:
				import pandas as pd
			except ImportError:
				messagebox.showerror(
					"Missing Library",
					"pandas is not installed.\n\nPlease run:\npip install pandas openpyxl"
				)
				return

			# Prepare data
			data = []
			for result in self.results:
				data.append({
					'File Name': result['filename'],
					'Type': result['file_type'],
					'Words': result['words'],
					'Cost (USD)': round(result['words'] * en_ko_rate, 2)
				})

			# Create DataFrame
			df = pd.DataFrame(data)

			# Add totals row
			totals = {
				'File Name': 'TOTAL',
				'Type': '',
				'Words': df['Words'].sum(),
				'Cost (USD)': df['Cost (USD)'].sum()
			}
			df = pd.concat([df, pd.DataFrame([totals])], ignore_index=True)

			# Generate filename with safe timestamp format
			timestamp = datetime.now().strftime('%Y%m%d-%H%M%S')
			output_file = f"gui_counter_report_{timestamp}.xlsx"

			# Export to Excel
			df.to_excel(output_file, index=False, sheet_name='Word Count')

			# Show success message with file location
			full_path = os.path.abspath(output_file)
			messagebox.showinfo(
				"Export Successful",
				f"Report saved successfully!\n\nFile: {output_file}\nLocation: {full_path}"
			)

			print(f"✓ Excel report saved: {output_file}")

		except Exception as e:
			# Show detailed error to user
			messagebox.showerror(
				"Export Failed",
				f"Failed to export Excel file.\n\nError: {str(e)}\n\nMake sure pandas and openpyxl are installed:\npip install pandas openpyxl"
			)
			print(f"Export error: {e}")

	# Run the application
if __name__ == "__main__":
	root = TkinterDnD.Tk()  # Use TkinterDnD for drag-and-drop
	app = WordCounterGUI(root)
	root.mainloop()



