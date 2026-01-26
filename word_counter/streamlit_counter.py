import streamlit as st
import os
import json
import xml.etree.ElementTree as ET
from docx import Document
import PyPDF2
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Word Counter", page_icon="📝", layout="wide")

st.title("📝 Localization Word Counter")
st.write("Upload your localization files to count words and estimate costs")

# File uploader
uploaded_files = st.file_uploader(
	"Choose files",
	type=['json', 'xml', 'xlf', 'docx', 'pdf'],
	accept_multiple_files=True
)


def count_words(file, file_type):
	"""Count words from uploaded file"""
	try:
		if file_type == 'json':
			data = json.load(file)
			# Extract strings logic here
			text = str(data)  # Simplified
		elif file_type in ['xml', 'xlf']:
			tree = ET.parse(file)
			text = ' '.join([elem.text for elem in tree.iter() if elem.text])
		elif file_type == 'docx':
			doc = Document(file)
			text = ' '.join([p.text for p in doc.paragraphs])
		elif file_type == 'pdf':
			pdf = PyPDF2.PdfReader(file)
			text = ' '.join([page.extract_text() for page in pdf.pages])

		return len(text.split())
	except:
		return 0


if uploaded_files:
	results = []

	with st.spinner("Processing files..."):
		for file in uploaded_files:
			ext = file.name.split('.')[-1].lower()
			words = count_words(file, ext)
			results.append({
				'File Name': file.name,
				'Type': ext.upper(),
				'Words': words,
				'Cost (USD)': round(words * 0.15, 2)
			})

	# Display results
	df = pd.DataFrame(results)

	st.success(f"✅ Processed {len(results)} files")

	col1, col2, col3 = st.columns(3)
	with col1:
		st.metric("Total Words", f"{df['Words'].sum():,}")
	with col2:
		st.metric("Estimated Cost", f"${df['Cost (USD)'].sum():,.2f}")
	with col3:
		st.metric("Files", len(results))

	st.dataframe(df, use_container_width=True)

	# Download button
	csv = df.to_csv(index=False)
	st.download_button(
		label="Download CSV",
		data=csv,
		file_name=f"streamlit_counter_report_{datetime.now().strftime('%Y%m%d')}.csv",
		mime="text/csv"
	)

