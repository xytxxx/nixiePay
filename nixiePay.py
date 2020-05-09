#coding=utf-8 
import json
import glob
import copy
import sys
import re
import os
import xlwings as xw
from collections import defaultdict

#=======================================================================================

listToInclude = '2020-04'
					# 把想要结算的列表的的名字放在引号里，所有名字包括引号内容的列表都会被结算
					# 例如：  如果引号内是'-'，项目是2019年1月，
					# 		 那么'2019-1' 列和'2018-12’列都会被结算

#=======================================================================================
#  以下为程序本体

errorTasks = {'有分段缺少结束时间': set(),
				'无法解析分段时间戳或没有时间戳': set(),
				'没有轴, 或者@的名字不对': set(),
				'分段时间戳与初翻检查项数量不符': set(),
				'没有校对, 或者@的名字不对': set(),
				'按D, P, S后加的名字找不到对应成员,请确认@对了人': set(),
				'未找到CNY任务(或者全员Paypal?)': set(),
				'有无校对分段但是D, P分段数量不一致': set()}

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
cards_all = {}
lists = {}
CNYmemberIds = []
cellStart = {'tasks': 'A3',
			'price': {'translator': 'J3', 'no_proofread': 'J4', 'proofreader': 'J5', 'subtitler': 'J6'},
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
		if listToInclude in lists[wekanCard['listId']] :
			cards[wekanCard['_id']] = {picked: wekanCard[picked] for picked in ["title", "createdAt", "description"]}		
			# initialize card data structure
			cards[wekanCard['_id']]['num_D_segments'] = 0
			cards[wekanCard['_id']]['num_P_segments'] = 0
			cards[wekanCard['_id']]['num_S_segments'] = 0
			cards[wekanCard['_id']]['num_D_segments_should_be'] = 0
			cards[wekanCard['_id']]['skip_proofread_segments'] = []
			cards[wekanCard['_id']]['id'] = wekanCard['_id']
			cards[wekanCard['_id']]['title_Bilibili'] = '未检测到'
			cards[wekanCard['_id']]['isClear'] = False
			cards[wekanCard['_id']]['error'] = False
			parseCardDescription(cards[wekanCard['_id']])


# inspect checklist items, give data to each team member
def parseChecklistItems(exportData):
	readChecklistItems = exportData['checklistItems'] 
	list_sorts = defaultdict(list)
	# count number of segments first 
	for wekanItem in readChecklistItems: 
		if wekanItem['cardId'] in cards:
			title = wekanItem['title']
			list_sorts[wekanItem['checklistId']].append(wekanItem['sort'])
			list_sorts[wekanItem['checklistId']] = sorted(list_sorts[wekanItem['checklistId']])
			if title[0] in ['D', 'P', 'S']:
				cards[wekanItem['cardId']]['num_'+title[0]+'_segments'] += 1
	# then count 无校对
	for wekanItem in readChecklistItems: 
		if wekanItem['cardId'] in cards:
			title = wekanItem['title']
			if '@免校对' in title:
				if cards[wekanItem['cardId']]['num_D_segments'] != cards[wekanItem['cardId']]['num_P_segments']:
					errorTasks['有无校对分段但是D, P分段数量不一致'].add(cards[wekanItem['cardId']]['title'])
					cards[wekanItem['cardId']]['error'] = True
				else:
					index = list_sorts[wekanItem['checklistId']].index(wekanItem['sort'])
					cards[wekanItem['cardId']]['skip_proofread_segments'].append(index)
	
	for wekanItem in readChecklistItems: 
		if wekanItem['cardId'] in cards:
			title = wekanItem['title']
			if title[0] in ['P', 'S']:
				# found a segment in cards
				member = wekanItem['title'].split('@', 1)[1]
				if member not in users:
					if '校对' in member:
						users['免校对'][title[0]].append(wekanItem['cardId'])
					else:
						errorTasks['按D, P, S后加的名字找不到对应成员,请确认@对了人'].add(cards[wekanItem['cardId']]['title'])
				elif not cards[wekanItem['cardId']]['error']:
					users[member][title[0]].append(wekanItem['cardId'])
			if title[0] == 'D':
				# found a segment in cards
				member = wekanItem['title'].split('@', 1)[1]
				if member not in users:
					errorTasks['按D, P, S后加的名字找不到对应成员,请确认@对了人'].add(cards[wekanItem['cardId']]['title'])
				elif not cards[wekanItem['cardId']]['error']:
					index = list_sorts[wekanItem['checklistId']].index(wekanItem['sort'])
					if index not in cards[wekanItem['cardId']]['skip_proofread_segments']:
						users[member]['D'].append(wekanItem['cardId'])
					else:
						users[member]['DP'].append(wekanItem['cardId'])
			if title[0] == 'T':
				cards[wekanItem['cardId']]['title_Bilibili'] = title[1:]
			

# initialize user directary		
def parseUserInfo(exportData):
	readUsers = exportData['users']
	for wekanUser in readUsers:
		users[wekanUser['username']] = {
			"D": [],
			"P": [],
			"DP": [],
			"S": [],
			"userName": wekanUser['username'],
			'id': wekanUser['_id']
		}
	users['免校对'] = {
		"D": [],
		"DP": [],
		"P": [],
		"S": [],
		"userName": '免校对',
		'id': '0'
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
			errorTasks['没有校对, 或者@的名字不对'].add(card['title'])
			foundError = True
		if card['num_S_segments'] is 0:
			errorTasks['没有轴, 或者@的名字不对'].add(card['title'])
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


#stolen from kilo19
def show_exception_and_exit(exc_type, exc_value, tb):
	printErrors()
	import traceback
	traceback.print_exception(exc_type, exc_value, tb)
	input('''脚本遇到错误，请截图此画面发送给王二麻。按回车键退出
	Error encountered, please send a screenshot of this error to DPang Er Ma
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
				   1 ]  # 1 为不拖欠，所有任务默认不拖欠
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

		if len(user['D']) + len(user['DP']) is not 0:
			duration_cells = []
			formula = '=' 
			for taskId in user['D']:
				if cards[taskId]['isClear']: 
					duration_cell = videoDuraionRow + str(clearTaskMapping[taskId])
					num_segments = str(cards[taskId]['num_D_segments'])
					duration_cells.append(duration_cell + '/' + num_segments)
			if len(duration_cells) > 0:
				duration_cells = sorted(duration_cells)
				formula += cellStart['price']['translator'] + '*' +  '(' + '+'.join(duration_cells) + ')'
			duration_cells = []
			for taskId in user['DP']:
				if cards[taskId]['isClear']: 
					duration_cell = videoDuraionRow + str(clearTaskMapping[taskId])
					num_segments = str(cards[taskId]['num_D_segments'])
					duration_cells.append(duration_cell + '/' + num_segments)
			if len(duration_cells) > 0:
				duration_cells = sorted(duration_cells)
				if formula != '=':
					formula += '+'
				formula += cellStart['price']['no_proofread'] + '*' +  '(' + '+'.join(duration_cells) + ')'
			if formula != '=':
				if cny is True:
					cnyTranslators.append([userName, formula])
				else:
					usdTranslators.append([userName, formula])
			
		if len(user['P']) is not 0:
			duration_cells = []
			for taskId in user['P']:
				if cards[taskId]['isClear']: 
					duration_cell = videoDuraionRow + str(clearTaskMapping[taskId])
					num_segments = str(cards[taskId]['num_P_segments'])
					duration_cells.append(duration_cell + '/' + num_segments)
				if len(duration_cells) > 0:
					duration_cells = sorted(duration_cells)
					formula = '=' + cellStart['price']['proofreader'] + '*' +  '(' + ' + '.join(duration_cells) + ')'
				else:
					formula = ''
			if formula != '':
				if cny is True:
					cnyProofreaders.append([user['userName'], formula])
				else:
					usdProofreaders.append([user['userName'], formula])

		if len(user['S']) is not 0:
			duration_cells = []
			for taskId in user['S']:
				if cards[taskId]['isClear']: 
					duration_cell = videoDuraionRow + str(clearTaskMapping[taskId])
					num_segments = str(cards[taskId]['num_S_segments'])
					duration_cells.append(duration_cell + '/' + num_segments)
				if len(duration_cells) > 0:
					duration_cells = sorted(duration_cells)
					formula = '=' +  cellStart['price']['subtitler'] + '*' +  '(' + '+'.join(duration_cells) + ')'
				else:
					formula = ''
			if formula != '':
				subtitlers.append([user['userName'], formula])
				
	

	xw.Range(cellStart['tally']['cnyTranslator']).value = cnyTranslators
	xw.Range(cellStart['tally']['usdTranslator']).value = usdTranslators
	xw.Range(cellStart['tally']['cnyProofreader']).value = cnyProofreaders 
	xw.Range(cellStart['tally']['usdProofreader']).value = usdProofreaders 
	xw.Range(cellStart['tally']['subtitler']).value = subtitlers 

def clearTally():
	scriptDir = os.path.dirname(os.path.realpath(sys.argv[0]))
	candidate = glob.glob(os.path.join(scriptDir, 'lmgns*.xlsx'))[0]
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
    
		
