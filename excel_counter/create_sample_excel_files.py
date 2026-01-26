"""
Create sample Excel files for localization word counter project
Includes tags in Korean source text to demonstrate tag stripping
"""

import pandas as pd
import os


def create_ui_strings():
	"""Create sample UI strings file with some tagged content"""
	data = {
		'String_ID': ['UI_001', 'UI_002', 'UI_003', 'UI_004', 'UI_005', 'UI_006',
					  'UI_007', 'UI_008', 'UI_009', 'UI_010', 'UI_011', 'UI_012',
					  'UI_013', 'UI_014', 'UI_015', 'UI_016', 'UI_017', 'UI_018',
					  'UI_019', 'UI_020'],
		'Context': ['Main Menu', 'Main Menu', 'Main Menu', 'Main Menu', 'HUD', 'HUD',
					'HUD', 'HUD', 'Inventory', 'Inventory', 'Inventory', 'Inventory',
					'Notification', 'Notification', 'Error', 'Error', 'Tutorial', 'Tutorial',
					'Quest', 'Quest'],
		'Korean': [
			'게임 시작',
			'계속하기',
			'<b>설정</b>',  # HTML tag
			'종료',
			'체력',
			'마나',
			'경험치',
			'레벨 {level}',  # placeholder
			'인벤토리',
			'무게: {current_weight}/{max_weight}',  # placeholders
			'장착',
			'사용',
			'<color=green>아이템을 획득했습니다</color>',  # Unity-style tag
			'레벨 업!',
			'인벤토리가 가득 찼습니다',
			'골드가 부족합니다',
			'WASD 키로 캐릭터를 이동하세요',
			'<i>스페이스바</i>를 눌러 공격하세요',  # HTML tag
			'마을 광장으로 가세요',
			'퀘스트 완료!'
		],
		'EN': [''] * 20,
		'JP': [''] * 20,
		'Character_Limit': [20, 20, 20, 15, 10, 10, 15, 10, 20, 10, 15, 15,
							30, 20, 40, 30, 60, 50, 40, 25],
		'Notes': ['Main menu button', 'Resume saved game', 'Options menu', 'Exit to desktop',
				  'Player health bar', 'Player mana bar', 'XP bar label', 'Character level',
				  'Inventory screen title', 'Current carry weight', 'Equip item button',
				  'Use item button', 'Item pickup message', 'Level up notification',
				  'Inventory full error', 'Insufficient currency', 'Movement tutorial',
				  'Combat tutorial', 'Quest objective', 'Quest completion']
	}

	df = pd.DataFrame(data)
	df.to_excel('ui_strings.xlsx', index=False, sheet_name='UI Strings')
	print("✓ Created: ui_strings.xlsx (with 5 tagged strings)")


def create_skill_descriptions():
	"""Create sample skill descriptions file with tagged content"""
	data = {
		'Skill_ID': ['SKL_001', 'SKL_002', 'SKL_003', 'SKL_004', 'SKL_005',
					 'SKL_006', 'SKL_007', 'SKL_008', 'SKL_009', 'SKL_010',
					 'SKL_011', 'SKL_012', 'SKL_013', 'SKL_014', 'SKL_015'],
		'Skill_Name': ['Fireball', 'Ice Shield', 'Lightning Strike', 'Healing Wave', 'Shadow Step',
					   'Poison Cloud', 'Divine Blessing', 'Earth Spike', 'Wind Dash', 'Meteor Storm',
					   'Stealth', 'Taunt', 'Resurrect', 'Berserk', 'Teleport'],
		'Type': ['Active', 'Defensive', 'Active', 'Support', 'Movement',
				 'DOT', 'Buff', 'Active', 'Movement', 'Ultimate',
				 'Utility', 'Control', 'Support', 'Buff', 'Movement'],
		'Korean': [
			'화염구를 발사하여 적에게 <b>큰 피해</b>를 줍니다.',  # HTML tag
			'얼음 방패를 생성하여 들어오는 피해를 {damage_reduction}% 감소시킵니다.',  # Placeholder
			'하늘에서 번개를 내리쳐 범위 피해를 줍니다.',
			'아군을 치유하고 체력을 회복시킵니다.',
			'그림자 속으로 순간이동하여 적의 공격을 회피합니다.',
			'독구름을 생성하여 지속 피해를 줍니다.',
			'신성한 축복으로 공격력과 방어력을 <color=gold>증가</color>시킵니다.',  # Unity tag
			'땅에서 바위 가시를 솟아나게 하여 적을 공격합니다.',
			'바람의 힘으로 빠르게 전방으로 돌진합니다.',
			'거대한 운석들을 소환하여 전장을 초토화시킵니다.',
			'은신 상태가 되어 적에게 발각되지 않습니다.',
			'적의 어그로를 끌어 자신을 공격하게 만듭니다.',
			'쓰러진 아군을 부활시킵니다.',
			'광폭화하여 공격 속도가 {speed_boost}% 증가하지만 방어력이 감소합니다.',  # Placeholder
			'지정한 위치로 즉시 이동합니다.'
		],
		'EN': [''] * 15,
		'JP': [''] * 15,
		'Cooldown': ['5s', '10s', '12s', '8s', '15s', '20s', '30s', '7s', '10s', '60s',
					 '25s', '12s', '180s', '45s', '20s'],
		'Notes': ['Basic fire spell', 'Damage reduction buff', 'AOE damage', 'Single target heal',
				  'Dodge ability', 'Area denial', 'Team buff', 'Ground-based attack',
				  'Gap closer', 'Ultimate ability', 'Invisibility', 'Tank ability',
				  'Revive spell', 'High risk buff', 'Long range teleport']
	}

	df = pd.DataFrame(data)
	df.to_excel('skill_descriptions.xlsx', index=False, sheet_name='Skills')
	print("✓ Created: skill_descriptions.xlsx (with 4 tagged strings)")


def create_dialogue():
	"""Create sample dialogue file - no tags for natural speech"""
	data = {
		'Dialogue_ID': ['DLG_001', 'DLG_002', 'DLG_003', 'DLG_004', 'DLG_005',
						'DLG_006', 'DLG_007', 'DLG_008', 'DLG_009', 'DLG_010',
						'DLG_011', 'DLG_012', 'DLG_013', 'DLG_014', 'DLG_015'],
		'Speaker': ['Hero', 'Elder', 'Elder', 'Hero', 'Elder',
					'Merchant', 'Merchant', 'Guard', 'Hero', 'Guard',
					'Villain', 'Villain', 'Hero', 'NPC', 'NPC'],
		'Chapter': ['Chapter 1', 'Chapter 1', 'Chapter 1', 'Chapter 1', 'Chapter 1',
					'Chapter 1', 'Chapter 1', 'Chapter 2', 'Chapter 2', 'Chapter 2',
					'Chapter 3', 'Chapter 3', 'Chapter 3', 'Side Quest', 'Side Quest'],
		'Korean': [
			'이 마을에 무슨 일이 일어난 거지? 모두가 사라졌어.',
			'젊은이여, 잘 왔네. 우리 마을이 위험에 처했다네.',
			'어둠의 세력이 우리 마을 사람들을 납치해 갔어. 자네의 도움이 필요하네.',
			'제가 어떻게 도와드릴 수 있죠?',
			'먼저 북쪽 숲으로 가서 잃어버린 고대 유물을 찾아주게.',
			'어서 오세요! 좋은 물건들이 많으니 구경해 보세요.',
			'이 검은 전설적인 대장장이가 만든 걸작이랍니다.',
			'여기서 더 이상 갈 수 없습니다. 통행증이 필요합니다.',
			'이 통행증으로 충분합니까?',
			'확인했습니다. 통과하셔도 됩니다. 조심하세요.',
			'드디어 만났군. 네가 소문의 그 영웅이냐?',
			'어리석은 자여, 네가 나를 막을 수 있다고 생각하나?',
			'당신의 악행은 여기서 끝입니다!',
			'제 고양이를 찾아주실 수 있나요? 집 뒤편에서 사라졌어요.',
			'정말 고마워요! 이 작은 선물을 받아주세요.'
		],
		'EN': [''] * 15,
		'JP': [''] * 15,
		'Type': ['Main Story', 'Main Story', 'Main Story', 'Main Story', 'Quest',
				 'Shop', 'Shop', 'Checkpoint', 'Checkpoint', 'Checkpoint',
				 'Boss Fight', 'Boss Fight', 'Boss Fight', 'Side Quest', 'Side Quest']
	}

	df = pd.DataFrame(data)
	df.to_excel('dialogue.xlsx', index=False, sheet_name='Dialogue')
	print("✓ Created: dialogue.xlsx (dialogue has no tags - natural speech)")


if __name__ == "__main__":
	print("\nCreating sample Excel files for localization project...")
	print("=" * 60)

	# Create samples folder
	if not os.path.exists('sample_excel_files'):
		os.makedirs('sample_excel_files')
		print("✓ Created folder: sample_excel_files/")

	os.chdir('sample_excel_files')

	create_ui_strings()
	create_skill_descriptions()
	create_dialogue()

	print("=" * 60)
	print("\n✓ All sample files created successfully!")
	print("\nFiles created:")
	print("  - ui_strings.xlsx")
	print("    - 5 strings with tags: <b>, {placeholders}, <color>")
	print("    - 15 strings without tags")
	print()
	print("  - skill_descriptions.xlsx")
	print("    - 4 strings with tags: <b>, {placeholders}, <color>")
	print("    - 11 strings without tags")
	print()
	print("  - dialogue.xlsx")
	print("    - 15 strings without tags (natural dialogue)")
	print()
	print("Tag types included:")
	print("  - HTML tags: <b></b>, <i></i>")
	print("  - Unity tags: <color=green></color>")
	print("  - Placeholders: {variable_name}")
	print()

