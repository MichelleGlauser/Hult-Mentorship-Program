from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
import os
import smtplib
import gspread

def get_matches(google_account, student_spreadsheet_url):
	matches_spreadsheet = google_account.open_by_url(student_spreadsheet_url)
	mentor_worksheet = matches_spreadsheet.get_worksheet(1)
	mentor_list_of_lists = mentor_worksheet.get_all_values()
	for item in mentor_list_of_lists:
		mentor_name = item[0]
		mentor_email = item[2]
		mentee_name1 = item[3] + " " + item[4]
		mentee_email1 = item[5]
		mentee_program1 = item[6]
		# mentorship_email = """\
		# 	<html>
		# 	<head></head>
		# 	<body>
		# 	<p>Dear """ + mentor_name + """,<br>
		# 	<p>Thank you for signing up to be a pioneer mentor in this brand-new Hult SF Alumni Mentorship Program! We are very excited to launch this program officially this Thursday, January 29th on the 4th floor of the Hult campus. The registration and check-in will begin at 7:30 PM on the 3rd floor. Please RSVP via Eventbrite using this link: <a href="tinyurl.com/ampmixer">AMP Mixer Event Page</a></p>
		# 	<p>Below is the contact information of your student mentee(s). We have encouraged them to reach out and connect with you prior to the kickoff event.</p><br>
		# 	<b>""" + mentee_name1 + " - " + mentee_email1 +  """</b><br>
		# 	<p>If you are unable to be at the event on Thursday, please let your mentee know. The time dedicated to this initiative will be decided between you and your mentees, so feel free to spark the conversation with them on the strategy that best suits you.</p>
		# 	<p>Please take a look at the <a href="hultsfamp.weebly.com">Alumni Mentorship Program page</a> prior to the event.
		# 	<p>With your support, we will greatly strengthen our Hult Alumni community in San Francisco and beyond. We look forward to seeing you at the kickoff party.</p>
		# 	<p>Many thanks,</p>
		# 	<p>San Francsico Hult Alumni Chapter</p><br>
		# 	<b>Walid E. Malouf</b><br>
		# 	Director, Career Services<br>
		# 	Hult San Francisco</p>
		# 	<img style="width:210px;height:34px;" src="http://rankings.r.ftdata.co.uk/lib/img/logos/entity/hult-international-business-school">
		# 	</body>
		# 	</html>
		# 	"""
		# if len(item) == 11:
		# 	mentee_name2 = item[6] + " " + item[7]
		# 	mentee_email2 = item[8]
		# 	mentee_program1 = item[6]
		# 	mentorship_email = """\
		# 		<html>
		# 		<head></head>
		# 		<body>
		# 		<p>Dear """ + mentor_name + """,<br>
		# 		<p>Thank you for signing up to be a pioneer mentor in this brand-new Hult SF Alumni Mentorship Program! We are very excited to launch this program officially this Thursday, January 29th on the 4th floor of the Hult campus. The registration and check-in will begin at 7:30 PM on the 3rd floor. Please RSVP via Eventbrite using this link: <a href="tinyurl.com/ampmixer">AMP Mixer Event Page</a></p>
		# 		<p>Below is the contact information of your student mentee(s). We have encouraged them to reach out and connect with you prior to the kickoff event.</p><br>
		# 		<b>""" + mentee_name1 + " - " + mentee_email1 + """</b><br>
		# 		<b>""" + mentee_name2 + " - " + mentee_email2 + """</b><br>
		# 		<b>""" + mentee_name1 + " - " + mentee_email1 + """</b><br>
		# 		<p>If you are unable to be at the event on Thursday, please let your mentee know. The time dedicated to this initiative will be decided between you and your mentees, so feel free to spark the conversation with them on the strategy that best suits you.</p>
		# 		<p>Please take a look at the <a href="hultsfamp.weebly.com">Alumni Mentorship Program page</a> prior to the event.
		# 		<p>With your support, we will greatly strengthen our Hult Alumni community in San Francisco and beyond. We look forward to seeing you at the kickoff party.</p>
		# 		<p>Many thanks,</p>
		# 		<p>San Francsico Hult Alumni Chapter</p><br>
		# 		<b>Walid E. Malouf</b><br>
		# 		Director, Career Services<br>
		# 		Hult San Francisco</p>
		# 		<img style="width:210px;height:34px;" src="http://rankings.r.ftdata.co.uk/lib/img/logos/entity/hult-international-business-school">
		# 		</body>
		# 		</html>
		# 		"""
		# if len(item) == 15:
		# 	mentee_name2 = item[6] + " " + item[7]
		# 	mentee_email2 = item[8]
		# 	mentee_name3 = item[9] + " " + item[10]
		# 	mentee_email3 = item[11]
		# 	mentee_program1 = item[6]
		# 	mentorship_email = """\
		# 		<html>
		# 		<head></head>
		# 		<body>
		# 		<p>Dear """ + mentor_name + """,<br>
		# 		<p>Thank you for signing up to be a pioneer mentor in this brand-new Hult SF Alumni Mentorship Program! We are very excited to launch this program officially this Thursday, January 29th on the 4th floor of the Hult campus. The registration and check-in will begin at 7:30 PM on the 3rd floor. Please RSVP via Eventbrite using this link: <a href="tinyurl.com/ampmixer">AMP Mixer Event Page</a></p>
		# 		<p>Below is the contact information of your student mentee(s). We have encouraged them to reach out and connect with you prior to the kickoff event.</p><br>
		# 		<b>""" + mentee_name1 + " - " + mentee_email1 + """</b><br>
		# 		<b>""" + mentee_name2 + " - " + mentee_email2 + """</b><br>
		# 		<b>""" + mentee_name3 + " - " + mentee_email3 + """</b><br>
		# 		<p>If you are unable to be at the event on Thursday, please let your mentee know. The time dedicated to this initiative will be decided between you and your mentees, so feel free to spark the conversation with them on the strategy that best suits you.</p>
		# 		<p>Please take a look at the <a href="hultsfamp.weebly.com">Alumni Mentorship Program page</a> prior to the event.
		# 		<p>With your support, we will greatly strengthen our Hult Alumni community in San Francisco and beyond. We look forward to seeing you at the kickoff party.</p>
		# 		<p>Many thanks,</p>
		# 		<p>San Francsico Hult Alumni Chapter</p><br>
		# 		<b>Walid E. Malouf</b><br>
		# 		Director, Career Services<br>
		# 		Hult San Francisco</p>
		# 		<img style="width:210px;height:34px;" src="http://rankings.r.ftdata.co.uk/lib/img/logos/entity/hult-international-business-school">
		# 		</body>
		# 		</html>
		# 		"""
		if len(item) == 19:
			mentee_name2 = item[7] + " " + item[8]
			mentee_email2 = item[9]
			mentee_program2 = item[10]
			print mentee_program1
			mentee_name3 = item[11] + " " + item[12]
			mentee_email3 = item[13]
			mentee_program3 = item[14]
			mentee_name4 = item[15] + " " + item[16]
			mentee_email4 = item[17]
			mentee_program4 = item[18]
			mentorship_email = """\
				<html>
				<head></head>
				<body>
				<p>Dear """ + mentor_name + """,<br>
				<p>Thank you for signing up to be a pioneer mentor in our brand-new Hult SF Alumni Mentorship Program! We are very excited to launch this program officially this Thursday, January 29th on the 4th floor of the Hult campus. The registration and check-in will begin at 7:30 PM on the 1st floor. Please RSVP via Eventbrite using this link: <a href="tinyurl.com/ampmixer">AMP Mixer Event Page</a></p>
				<p>Below is the contact information of your student mentee(s). We have encouraged them to reach out and connect with you prior to the kickoff mixer.</p>
				<b>""" + mentee_name1 + " " + mentee_program1 + " " + mentee_email1 + """</b><br>
				<b>""" + mentee_name2 + " " + mentee_program2 + " " + mentee_email2 + """</b><br>
				<b>""" + mentee_name3 + " " + mentee_program3 + " " + mentee_email3 + """</b><br>
				<b>""" + mentee_name4 + " " + mentee_program4 + " " + mentee_email4 + """</b><br>
				<p>If you are unable to be at the event on Thursday, please let your mentee(s) know. You and your mentee(s) should decide the amount of time to dedicate to this program. Please take a look at the <a href="hultsfamp.weebly.com">Alumni Mentorship Program page</a> prior to the event for more information.
				<p>With your support, we will greatly strengthen our Hult Alumni community in San Francisco and beyond. We look forward to seeing you at the kickoff mixer!</p>
				<p>Many thanks,</p>
				<p><b>Jessica Loman</b><br>
				President, SF Alumni Chapter
				<p><b>Walid E. Malouf</b><br>
				Director, Career Services<br>
				Hult San Francisco</p>
				<img style="width:210px;height:34px;" src="http://rankings.r.ftdata.co.uk/lib/img/logos/entity/hult-international-business-school">
				</body>
				</html>
				"""
			print mentor_email, mentorship_email
			send_email(mentorship_email, mentor_email)

def send_email(mentorship_email, mentor_email):
	hult_email = os.environ['hult_email']
	ef_email = os.environ['ef_email']
	hult_pass = os.environ['hult_pass']
	ccaddr = os.environ['ccaddr']
	toaddr = [mentor_email, ccaddr]
	msg = MIMEMultipart()
	msg['From'] = hult_email
	msg['To'] = ','.join(toaddr)
	msg['Subject'] = "Connect With Your Hult Mentee(s)!"
	msg.attach(MIMEText(mentorship_email, 'html'))

	server = smtplib.SMTP('outlook.office365.com', 587)
	server.ehlo()
	server.starttls()
	server.ehlo()
	server.login(ef_email, hult_pass)
	text = msg.as_string()
	server.sendmail(hult_email, toaddr, text)

def main():
	google_account = gspread.login(os.environ['gmail'], os.environ['password'])
	student_spreadsheet_url = os.environ['student_url']
	students = get_matches(google_account, student_spreadsheet_url)

if __name__ == '__main__':
    main()
