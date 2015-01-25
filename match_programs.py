import os
import gspread

def get_alumni(google_account, alumni_spreadsheet_url):
	alumni_spreadsheet = google_account.open_by_url(alumni_spreadsheet_url)
	alumni_worksheet = alumni_spreadsheet.get_worksheet(1)
	alumni_list_of_lists = alumni_worksheet.get_all_values()
	alumni = []
	for list in alumni_list_of_lists:
		alum = {}
		alum['name'] = list[0] + " " + list[1]
		alum['program'] = list[2]
		alum['region'] = list[3]
		alum['keywords'] = list[7].split(', ')
		alum['nationality'] = list[8]
		alumni.append(alum)
	return alumni

def get_students(google_account, student_spreadsheet_url):
	student_spreadsheet = google_account.open_by_url(student_spreadsheet_url)
	student_worksheet = student_spreadsheet.get_worksheet(1)
	student_list_of_lists = student_worksheet.get_all_values()
	students = []
	for list in student_list_of_lists:
		student = {}
		student['name'] = list[0] + " " + list[1]
		student['id'] = list[2]
		student['program'] = list[3]
		student['region'] = list[4]
		student['keywords'] = list[5].split(', ')
		student['nationality'] = list[8]
		student['mentor'] = []
		students.append(student)
	return students

def match_up(students, alumni):
	for student in students:
		for alum in alumni:
			if student['program'] == alum['program']:
				region_score = 1 if student['region'] == alum['region'] else 0
				nationality_score = 1 if student['nationality'] == alum['nationality'] else 0
				matching_keywords = set(student['keywords']) & set(alum['keywords'])
				keywords_score = len(matching_keywords)
				scores = (region_score, nationality_score, matching_keywords, keywords_score)
				sum_scores = keywords_score + nationality_score + region_score
				student['mentor'].append(["Program matches.", sum_scores, scores, alum['name'], alum['nationality']])
	print students

def main():
	google_account = gspread.login(os.environ['gmail'], os.environ['password'])
	alumni_spreadsheet_url = os.environ['alumni_url']
	student_spreadsheet_url = os.environ['student_url']
	alumni = get_alumni(google_account, alumni_spreadsheet_url)
	students = get_students(google_account, student_spreadsheet_url)
	matches = match_up(students, alumni)

if __name__ == '__main__':
    main()
