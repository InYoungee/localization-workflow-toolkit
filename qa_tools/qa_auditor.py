import json
import re
import pandas as pd

def extract_placeholders(text):
	"""Find all {tags} and %d/%s placeholders"""
	pattern = r"\{.*?\}|%\w"
	return re.findall(pattern, text)

def extract_html_tags(text):
	pattern = r"<[^>]*>"
	return re.findall(pattern, text)

def run_qa_audit(source_file, target_file):
	#Load the JSON data
	with open(source_file, 'r', encoding='utf-8') as f:
		source_data = json.load(f)
	with open(target_file, 'r', encoding='utf-8') as f:
		target_data = json.load(f)

	report_list = []

	for key, source_text in source_data.items():
		target_text = target_data.get(key, "")

		# 1. Placeholder check
		s_ph = extract_placeholders(source_text)
		t_ph = extract_placeholders(target_text)
		if s_ph != t_ph:
			report_list.append({
				"Key": key,
				"Issue": "Placeholder Missing",
				"Severity": "High",
				"Source": source_text,
				"Target": target_text
			})

		# 2. HTML Integrity check
		s_tags = extract_html_tags(source_text)
		t_tags = extract_html_tags(target_text)
		if s_tags != t_tags:
			report_list.append({
				"Key": key,
				"Issue": "HTML Tag Corruption",
				"Severity": "CRITICAL",
				"Source": source_text,
				"Target": target_text
			})

		# 3. Length/Expansion check (Flag if target is > *2 source length
		if len(target_text) > (len(source_text) * 2):
			report_list.append({
				"Key": key,
				"Issue": "Potential UI Overflow",
				"Severity": "Warning",
				"Source": source_text,
				"Target": target_text
			})

	df = pd.DataFrame(report_list)
	if not df.empty:
		df.to_excel("Localization_QA_Report.xlsx", index=False)
		print(f"Report generated with {len(df)} issues.")
	else:
		print("No issues found!")

run_qa_audit('qa_en-US.json', 'qa_ko-KR.json')