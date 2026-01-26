"""
Excel column word counter with tag stripping functionality
"""

import pandas as pd
import os
import glob
import re
from datetime import datetime


class ExcelColumnCounter:
	"""Extract and count words from specific Excel columns, stripping tags"""

	def __init__(self, target_columns=None, strip_tags=True):
		"""
		Initialize with target column names

		Args:
			target_columns: List of column names to extract text from
			strip_tags: Whether to strip HTML/Unity tags and placeholders
		"""
		if target_columns is None:
			self.target_columns = ['Korean', 'KO', 'Source', 'Source Text', 'korean']
		else:
			self.target_columns = target_columns

		self.strip_tags = strip_tags
		self.results = []

	def clean_text(self, text):
		"""
		Remove tags and placeholders from text

		Args:
			text: String to clean

		Returns:
			Cleaned text with tags removed
		"""
		if pd.isna(text):
			return ""

		text = str(text)

		if not self.strip_tags:
			return text

		# Remove HTML tags: <b>text</b>, <i>text</i>
		text = re.sub(r'<[^>]+>', '', text)

		# Remove Unity-style tags: <color=red>text</color>
		text = re.sub(r'<[^>]+>[^<]*</[^>]+>', lambda m: re.sub(r'<[^>]+>|</[^>]+>', '', m.group()), text)

		# Remove placeholders: {variable}, {{name}}
		text = re.sub(r'\{[^}]+\}', '', text)

		# Remove extra whitespace
		text = ' '.join(text.split())

		return text

	def find_text_column(self, df):
		"""Find which column contains the translatable text"""
		# Check exact matches first
		for col in self.target_columns:
			if col in df.columns:
				return col

		# Check case-insensitive matches
		df_columns_lower = {col.lower(): col for col in df.columns}
		for target in self.target_columns:
			if target.lower() in df_columns_lower:
				return df_columns_lower[target.lower()]

		return None

	def count_words_in_file(self, filepath, column_name=None):
		"""Count words in a specific column of an Excel file"""
		filename = os.path.basename(filepath)

		try:
			df = pd.read_excel(filepath)

			# Find the text column
			if column_name:
				if column_name not in df.columns:
					print(f"  ✗ Column '{column_name}' not found in {filename}")
					return None
				text_column = column_name
			else:
				text_column = self.find_text_column(df)
				if not text_column:
					print(f"  ✗ No text column found in {filename}")
					return None

			# Extract text from the column
			text_data = df[text_column].dropna()

			# Count words with and without tags
			words_with_tags = sum(len(str(text).split()) for text in text_data)

			# Clean tags and count again
			cleaned_texts = [self.clean_text(text) for text in text_data]
			words_without_tags = sum(len(text.split()) for text in cleaned_texts if text)

			# Count how many strings had tags
			strings_with_tags = sum(1 for original, cleaned in zip(text_data, cleaned_texts)
									if str(original) != cleaned)

			result = {
				'filename': filename,
				'column_used': text_column,
				'rows': len(text_data),
				'words_with_tags': words_with_tags,
				'words_without_tags': words_without_tags,
				'strings_with_tags': strings_with_tags
			}

			print(f"  ✓ {filename}")
			print(f"    Column: '{text_column}'")
			print(f"    Rows: {len(text_data):,}")
			print(f"    Words (with tags): {words_with_tags:,}")
			print(f"    Words (clean): {words_without_tags:,}")
			print(f"    Strings with tags: {strings_with_tags}")

			return result

		except Exception as e:
			print(f"  ✗ Error processing {filename}: {e}")
			return None

	def count_folder(self, folder_path, column_name=None):
		"""Count words in all Excel files in a folder"""
		excel_files = []
		for pattern in ['*.xlsx', '*.xls']:
			excel_files.extend(glob.glob(os.path.join(folder_path, pattern)))

		if not excel_files:
			print(f"No Excel files found in {folder_path}")
			return

		print(f"\nFound {len(excel_files)} Excel file(s)\n")
		print("=" * 70)

		self.results = []

		for filepath in excel_files:
			result = self.count_words_in_file(filepath, column_name)
			if result:
				self.results.append(result)
			print()

	def display_summary(self):
		"""Display summary of results"""
		if not self.results:
			print("No results to display.")
			return

		print("=" * 80)
		print("SUMMARY (Tag Stripping Enabled)" if self.strip_tags else "SUMMARY")
		print("=" * 80)
		print(f"{'File':<30} {'Column':<12} {'Rows':>6} {'Words':>10} {'Tagged':>8}")
		print("-" * 80)

		total_rows = 0
		total_words = 0
		total_tagged = 0

		for result in self.results:
			words = result['words_without_tags'] if self.strip_tags else result['words_with_tags']
			print(f"{result['filename']:<30} "
				  f"{result['column_used']:<12} "
				  f"{result['rows']:>6,} "
				  f"{words:>10,} "
				  f"{result['strings_with_tags']:>8}")
			total_rows += result['rows']
			total_words += words
			total_tagged += result['strings_with_tags']

		print("-" * 80)
		print(f"{'TOTAL':<30} {'':<12} {total_rows:>6,} {total_words:>10,} {total_tagged:>8}")
		print("=" * 80)

		cost_per_word = 0.15
		estimated_cost = total_words * cost_per_word

		print(f"\nTotal strings: {total_rows:,}")
		print(f"Strings with tags: {total_tagged}")
		print(f"Total words (clean): {total_words:,}")
		print(f"Estimated cost: ${estimated_cost:,.2f} (at ${cost_per_word}/word)")
		print("=" * 80)

	def export_to_excel(self, output_filename=None):
		"""Export results to Excel report"""
		if not self.results:
			print("No results to export.")
			return

		try:
			# Prepare data
			data = []
			for result in self.results:
				words = result['words_without_tags'] if self.strip_tags else result['words_with_tags']
				data.append({
					'File Name': result['filename'],
					'Column': result['column_used'],
					'Rows': result['rows'],
					'Words (Clean)': words,
					'Strings with Tags': result['strings_with_tags'],
					'Cost (USD)': round(words * 0.15, 2)
				})

			df = pd.DataFrame(data)

			# Add totals
			totals = {
				'File Name': 'TOTAL',
				'Column': '',
				'Rows': df['Rows'].sum(),
				'Words (Clean)': df['Words (Clean)'].sum(),
				'Strings with Tags': df['Strings with Tags'].sum(),
				'Cost (USD)': df['Cost (USD)'].sum()
			}
			df = pd.concat([df, pd.DataFrame([totals])], ignore_index=True)

			if not output_filename:
				timestamp = datetime.now().strftime('%Y%m%d-%H%M%S')
				output_filename = f"excel_word_count_{timestamp}.xlsx"

			df.to_excel(output_filename, index=False, sheet_name='Word Count')

			print(f"\n✓ Report exported: {output_filename}")
			return output_filename

		except Exception as e:
			print(f"✗ Export failed: {e}")
			return None


if __name__ == "__main__":
	print("\n" + "=" * 70)
	print("EXCEL COLUMN WORD COUNTER (WITH TAG STRIPPING)")
	print("=" * 70)

	# Count with tag stripping enabled
	counter = ExcelColumnCounter(strip_tags=True)
	counter.count_folder('./sample_excel_files')
	counter.display_summary()
	counter.export_to_excel()
