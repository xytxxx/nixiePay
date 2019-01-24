#coding=utf-8 
import json
import glob
import copy
import sys
import re
import os
import time	
import xlwings as xw

listToInclude = '2018-12' #change this to the title of the list

errorTasks = {'有分段缺少结束时间': set(),
				'无法解析分段时间戳或没有时间戳': set(),
				'没有轴': set(),
				# '初翻检查项没有翻译的名字': set(),
				'分段时间戳与初翻检查项数量不符': set(),
				# '校对检查项没有二校名字或名字不是二校\n如果是二校负责了本任务二校并且检查项无误，请联系Kilo19': set(),
				# '检查项错误\n必须D、P、T三者之一开头（代表初翻、二校、标题），后面跟半角空格\nD与P项需要有@人名，@与D和P后的半角空格之间可以有任何字符': set(),
				'没有校对': set(),
				'按D, P, S后加的名字找不到对应成员': set(),
				'未找到CNY任务(或者全员Paypal?)': set()}

# cards has the format:
# {
#     "keuXvBnmwS2ZoFrxm": {
#         "title": "LTT@CES: This TV changes EVERYTHING - HiSense",
#         "createdAt": "2019-01-14T19:42:17.852Z",
#         "description": "//00:00-03:41 @初翻\n//03:41-07:22 @初翻\n\n——————————————\ngoogl...",
#		  "duration": 723 (in seconds)
# 		  "num_D_segments": 2,
# 		  "num_D_segments_should_be": 2
# 		  "num_P_segments": 1,
# 		  "num_S_segments": 2,
#		  "title_Bilibili": "LTT在CES展: 海信改变一切"
#     }
# }


# lists has the format:
# {
#     'OxcqweizBVhj12jJH': "公用",
#     ...
# }


# users has the format:
# {
# 	"王二麻"： {
# 		"D": ["keuXvBnmwS2ZoFrxm", "asdas211deqc123", ...] (id's of cards)
# 		"P"：[...],
# 		"S": [...],
#		"userName": “王二麻”,
#       "id": "12b4id8aklad8"
# 	}, 
# 	"kilo19": {
# 		...
# 	}
# 	...
# }
cards = {}
users = {}
lists = {}
CNYmemberIds = []
cellStart = {'tasks': 'A3',
			'price': {'translator': 'J3', 'proofreader': 'J4', 'subtitler': 'J5'},
			'tally': {'cnyTranslator': 'E3', 'usdTranslator': 'G3',
					  'cnyProofreader': 'I21', 'usdProofreader': 'K21',
					  'subtitler': 'I12'}}

	
# stolen from kilo19
def parseCardDescription(card):
	minute = 0
	sec = 0
	try: 
		allTimes = re.findall(r'\d{1,2}:\d{1,2}', card['description'])  
		if len(allTimes) % 2:
			errorTasks['有分段缺少结束时间'].add(card['title'])
			card["num_D_segments_should_be"] = 0
		elif len(allTimes) >= 2:
			min = int(allTimes[-1][:2]) - int(allTimes[0][:2])
			sec = int(allTimes[-1][3:]) - int(allTimes[0][3:])
			card["num_D_segments_should_be"] = len(allTimes) / 2
		else:
			errorTasks['无法解析分段时间戳或没有时间戳'].add(card['title'])
			card["num_D_segments_should_be"] = 0
	except ValueError:
		errorTasks['无法解析分段时间戳或没有时间戳'].add(card['title'])
		card['duration'] = 0
		card["num_D_segments_should_be"] = 0
	else:
		card['duration'] = min * 60 + sec

# build list id-title lookup table
def parseListsInfo(exportData):
	readLists = exportData['lists']
	for wekanList in readLists: 
		lists[wekanList['_id']] = wekanList['title']

# populate cards
def parseCardInfo(exportData):
	readCards = exportData['cards']
	for wekanCard in readCards:
		if wekanCard['title'] == 'CNY' and len(CNYmemberIds) == 0:
			CNYmemberIds[:] = wekanCard['members'][:]
		# only include cards of desired list
		if lists[wekanCard['listId']] == listToInclude:
			cards[wekanCard['_id']] = {picked: wekanCard[picked] for picked in ["title", "createdAt", "description"]}		
			# initialize card data structure
			cards[wekanCard['_id']]['num_D_segments'] = 0
			cards[wekanCard['_id']]['num_P_segments'] = 0
			cards[wekanCard['_id']]['num_S_segments'] = 0
			cards[wekanCard['_id']]['num_D_segments_should_be'] = 0
			cards[wekanCard['_id']]['id'] = wekanCard['_id']
			parseCardDescription(cards[wekanCard['_id']])
		


# inspect checklist items, give data to each team member
def parseChecklistItems(exportData):
	readChecklistItems = exportData['checklistItems'] 
	for wekanItem in readChecklistItems: 
		if wekanItem['cardId'] in cards:
			title = wekanItem['title']
			if title[0] in ['D', 'P', 'S']:
				# found a segment in cards
				member = wekanItem['title'].split('@', 1)[1]
				if member not in users:
					errorTasks['按D, P, S后加的名字找不到对应成员'].add(cards[wekanItem['cardId']]['title'])
				else: 
					cards[wekanItem['cardId']]['num_'+title[0]+'_segments'] += 1
					users[member][title[0]].append(wekanItem['cardId'])
			if title[0] == 'T':
				cards[wekanItem['cardId']]['title_Bilibili'] = title[1:]
			
			

# initialize user directary		
def parseUserInfo(exportData):
	readUsers = exportData['users']
	for wekanUser in readUsers:
		users[wekanUser['username']] = {
			"D": [],
			"P": [],
			"S": [],
			"userName": wekanUser['username'],
			'id': wekanUser['_id']
		}

def validateCards():
	if len(CNYmemberIds) is 0:
		errorTasks['未找到CNY任务(或者全员Paypal?)'].add(0)
	for cardId, card in cards.items():
		foundError = False
		if ('num_D_segments_should_be' not in card) or (int(card ['num_D_segments']) != int(card['num_D_segments_should_be'])):
			errorTasks['分段时间戳与初翻检查项数量不符'].add(card['title'])
			foundError = True
		if card['num_P_segments'] is 0:
			errorTasks['没有校对'].add(card['title'])
			foundError = True
		if card['num_S_segments'] is 0:
			errorTasks['没有轴'].add(card['title'])
			foundError = True
		if not foundError:
			card['isClear'] = True

def printErrors():
	for key, values in errorTasks.items():
		if values:
			print (key)
			for v in values:
				print(v)
			print()

def show_exception_and_exit(exc_type, exc_value, tb):
	printErrors()
	import traceback
	traceback.print_exception(exc_type, exc_value, tb)
	input('''脚本遇到错误，请截图此画面发送给王二麻。按回车键退出
	Error encountered, please send a screenshot of this error to Wang Er Ma
	Press Enter to Exit
	''')
	sys.exit(-1)

sys.excepthook = show_exception_and_exit

#stolen from kilo19
def writeTasks():
	rows = []
	clearTaskMapping = {}
	tasks = sorted(cards.values(), key=lambda card: card['createdAt']) 
	for x in range(len(tasks)):
		task = tasks[x]
		if 'isClear' in task:
			curRowOffset = len(rows)
			clearTaskMapping[task['id']] = curRowOffset + int(cellStart['tasks'][1:])
			row = [task['title'],
				   #精确到.25分钟
				   round(task['duration'] / 15) / 4,
				   1 ]  # 1 意为不拖欠，所有任务默认不拖欠
			rows.append(row)
	xw.Range(cellStart['tasks']).value = rows
	return clearTaskMapping
 
def writeSalary(clearTaskMapping):
	cnyTranslators = []
	usdTranslators = []
	cnyProofreaders = []
	usdProofreaders = []
	subtitlers = []
	videoDuraionRow = chr(ord(cellStart['tasks'][0]) + 3)   
	for userName, user in users.items():
		cny = True
		if user['id'] in CNYmemberIds:
			cny = True
		else: 
			cny = False

		if len(user['D']) is not 0:
			duration_cells = []
			for taskId in user['D']:
				duration_cell = videoDuraionRow + str(clearTaskMapping[taskId])
				num_segments = str(cards[taskId]['num_D_segments'])
				duration_cells.append(duration_cell + '/' + num_segments)
			formula = '=' + cellStart['price']['translator'] + '*' +  '(' + '+'.join(duration_cells) + ')'
			if cny is True:
				cnyTranslators.append([userName, formula])
			else:
				usdTranslators.append([userName, formula])
		
		if len(user['P']) is not 0:
			duration_cells = []
			for taskId in user['P']:
				duration_cell = videoDuraionRow + str(clearTaskMapping[taskId])
				num_segments = str(cards[taskId]['num_P_segments'])
				duration_cells.append(duration_cell + '/' + num_segments)
			formula = '=' + cellStart['price']['proofreader'] + '*' +  '(' + ' + '.join(duration_cells) + ')'
			if cny is True:
				cnyProofreaders.append([user['userName'], formula])
			else:
				usdProofreaders.append([user['userName'], formula])

		if len(user['S']) is not 0:
			duration_cells = []
			for taskId in user['S']:
				duration_cell = videoDuraionRow + str(clearTaskMapping[taskId])
				num_segments = str(cards[taskId]['num_S_segments'])
				duration_cells.append(duration_cell + '/' + num_segments)
			formula = '=' +  cellStart['price']['subtitler'] + '*' +  '(' + '+'.join(duration_cells) + ')'
			subtitlers.append([user['userName'], formula])
	

	xw.Range(cellStart['tally']['cnyTranslator']).value = cnyTranslators
	time.sleep(.300)
	xw.Range(cellStart['tally']['usdTranslator']).value = usdTranslators
	time.sleep(.300)
	xw.Range(cellStart['tally']['cnyProofreader']).value = cnyProofreaders 
	time.sleep(.300)
	xw.Range(cellStart['tally']['usdProofreader']).value = usdProofreaders 
	time.sleep(.300)
	xw.Range(cellStart['tally']['subtitler']).value = subtitlers 

def clearTally():
	scriptDir = os.path.dirname(os.path.realpath(sys.argv[0]))
	candidate = glob.glob(os.path.join(scriptDir, 'LMGNS*.xlsx'))[0]
	wb = xw.Book(candidate)
	clearTaskMapping = writeTasks()
	writeSalary(clearTaskMapping)


# read json file
def main():
		

	scriptDir = os.path.dirname(os.path.realpath(sys.argv[0]))
	candidates = glob.glob(os.path.join(scriptDir, 'wekan-export-*'))
	for candidate in candidates:
		inf = open(candidate, encoding = 'utf-8', mode = 'r')
		readData = json.load(inf)

		parseListsInfo(readData)

		parseUserInfo(readData)
		
		parseCardInfo(readData)

		parseChecklistItems(readData)

		validateCards()
	
	clearTally()
	printErrors()

	input('Clear!!')


if __name__ == "__main__":
	main()
    
		
