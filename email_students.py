from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
import os
import smtplib
import gspread

def get_matches(google_account, student_spreadsheet_url):
	student_spreadsheet = google_account.open_by_url(student_spreadsheet_url)
	student_worksheet = student_spreadsheet.get_worksheet(0)
	student_list_of_lists = student_worksheet.get_all_values()
	for item in student_list_of_lists:
		student_name = item[0]
		mentor_name = item[2]
		mentor_company = item[3]
		student_email = item[4]
		mentor_email = item[5]
		send_email(student_name, mentor_name, mentor_company, student_email, mentor_email)

def send_email(student_name, mentor_name, mentor_company, student_email, mentor_email):
	hult_email = os.environ['hult_email']
	ef_email = os.environ['ef_email']
	hult_pass = os.environ['hult_pass']
	fromaddr = hult_email
	ccaddr = os.environ['ccaddr']
	toaddr = [student_email, ccaddr]

	msg = MIMEMultipart()
	msg['From'] = fromaddr
	msg['To'] = ','.join(toaddr)
	msg['Subject'] = "Connect With Your Hult Alumni Mentor!"

	mentorship_email = """\
	<html>
	<head></head>
	<body>
	<p>Dear """ + student_name + """,<br>
	<p>Thank you for signing up for the Hult SF Alumni Mentorship Program. We are very excited to launch this program officially this Thursday, January 29th on the 4th floor of the campus. The registration and check-in will begin at 7:30 PM on the 3rd floor. Please RSVP via Eventbrite using this link: <a href="tinyurl.com/ampmixer">AMP Mixer Event Page</a></p>
	<p><b>Your alumni mentor is """ + mentor_name + " - " + mentor_company + " - " + mentor_email + """</b><br>
	<p>We highly encourage you to reach out and connect with your mentor prior to the kickoff mixer. If you are unable to be there on Thursday, please let your mentor know. You and your mentor should decide the amount of time to dedicate to this program.</p>
	<p>Please take a look at the <a href="hultsfamp.weebly.com">Alumni Mentorship Program page</a> prior to the event for more information.
	<p>We look forward to seeing you at the kickoff mixer!</p>
	<p>Thanks,</p>
	<p><b>Walid E. Malouf</b><br>
	Director, Career Services<br>
	Hult San Francisco</p>
	<img style="width:210px;height:34px;" src="http://rankings.r.ftdata.co.uk/lib/img/logos/entity/hult-international-business-school">
	</body>
	</html>
	"""
	print mentorship_email

	msg.attach(MIMEText(mentorship_email, 'html'))

	server = smtplib.SMTP('outlook.office365.com', 587)
	server.ehlo()
	server.starttls()
	server.ehlo()
	server.login(ef_email, hult_pass)
	text = msg.as_string()
	server.sendmail(fromaddr, toaddr, text)

def main():
	google_account = gspread.login(os.environ['gmail'], os.environ['password'])
	student_spreadsheet_url = os.environ['student_url']
	students = get_matches(google_account, student_spreadsheet_url)

if __name__ == '__main__':
    main()
