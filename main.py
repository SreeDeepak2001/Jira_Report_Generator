import datetime
import json
import re
from xml.dom.minidom import parse
from jira import JIRA
import win32com.client as client
from requests.auth import HTTPBasicAuth
from html_body import HTML_BODY
import requests
import logging
import logging.handlers

api_token = ""
email = ""
server = "https://Servername.atlassian.net"

logger = logging.getLogger()
logger.setLevel(logging.INFO)
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
handler = logging.handlers.RotatingFileHandler(filename='JIRAReminderTool.log', mode='w', backupCount=1)
handler.setLevel(logging.INFO)
handler.setFormatter(formatter)
logger.addHandler(handler)


class JiraTool:

    def __init__(self):
        self.jira = JIRA(basic_auth=(email, api_token), server=server)
        logger.info('Created Jira Object')
        try:
            logger.info('Parsing ProjectID.xml')
            doc = parse("ProjectID.xml")  # XML where the project IDs are stored
        except Exception as e:
            logger.error(f'Cannot parse ProjectID.xml. Exception : {e}')
            exit()
        logger.info('ProjectID.xml parsed')
        projects = doc.getElementsByTagName("project")
        logger.info('Retrieving all project from ProjectID.xml')
        for project in projects:
            log = self.getnode(project, 'logger')
            print(log.upper())
            if log == 'None':
                self.getlogger('ERROR')
            else:
                self.getlogger(log.upper())
            Id = project.getElementsByTagName("ID")
            if Id[0].firstChild is None:
                logger.info('No Id for the project')
            else:
                logger.info(f'Creating CONFIG for project : {Id[0].firstChild.nodeValue}(ID)')
                confif = self.projectconfig(project)
                logger.info(f'Validating the CONFIG for project : {Id[0].firstChild.nodeValue}(ID)')
                if self.validateConfig(confif) == 1:
                    logger.info('CONFIG valid')
                    self.projectcheck(confif)
                    pass
                else:
                    logger.error('Invalid CONFIG')
                    logger.error('Exiting Program')
                    exit()

    def projectconfig(self, project):
        """
        Creating a CONFIG Dict for the project to search in JIRA
        :param project: project node in the ProjectID.xml file (node object)
        :return: CONFIG dictionary for the project (dict)
        """
        projectID = self.getnode(project, "ID")
        P1 = self.getnode(project, "Critical")
        P2 = self.getnode(project, "Severe")
        P3 = self.getnode(project, "Moderate")
        P4 = self.getnode(project, "Minor")
        moreInfo = self.getnode(project, "MoreInformation")
        commentConfig = self.getnode(project, "comment_config").lower()
        sender = project.getElementsByTagName("mailID")
        mailID = ""
        for mail in sender:
            if mail.firstChild is not None:
                mailID = mailID + mail.firstChild.nodeValue + ';'
        if mailID == "":
            mailID = "None"
        CONFIG = {
            "Project ID": projectID,
            "1 - Critical": P1,
            "2 - Severe": P2,
            "3 - Moderate": P3,
            "4 - Minor": P4,
            "More Information": moreInfo,
            "Comment Type": commentConfig,
            "sender": mailID
        }
        return CONFIG

        pass

    def projectcheck(self, CONFIG):
        """
            To Gather the tickets and their assignee categorized by the project code,
             and get the ticket based on their CONFIG.
             :param CONFIG: CONFIG Dictionary for the project (dict)
        """
        projectID = CONFIG.get("Project ID")
        mailID = CONFIG.get("sender")
        jquery = f"project = {projectID} AND status not in (Closed,Close,Complete,Completed) ORDER BY priority DESC"
        logger.info(f'sending query \" {jquery} \" to jira')
        issueCount = 0
        self.tickets = []
        startat = 0
        logger.info(f'Retrieving issues of project : {projectID}(ID)')
        try:
            while True:
                issues = self.jira.search_issues(jql_str=jquery, maxResults=1000, startAt=startat)
                self.addticketlist(issues, CONFIG)
                startat = startat + 100
                issueCount = issueCount + len(issues)
                issues = self.jira.search_issues(jql_str=jquery, maxResults=1000, startAt=startat)
                if len(issues) == 0:
                    break
            logger.info(f'Total tickets received {issueCount}')
            tableContent = self.gettable(self.tickets)
            if len(tableContent) != 0:
                self.sendmail(mailID, tableContent)
            else:
                logger.info('No tickets need attention')
                pass
        except Exception as e:
            logger.error(f'Cannot retrieve issues for project : {projectID}(ID). Facing exception \"{e.args[0]}\"')
            logger.error('Exiting Program')

    def addticketlist(self, issues, CONFIG):
        """
            Add sorted tickets based on the CONFIG to a LIST
        :param issues: List of issues from JIRA API (list)
        :param CONFIG: CONFIG Dictionary for the project (dict)
        :return: List of Tickets. where, Ticket are the Dict of issue info formed based on the CONFIG given (list)
        """
        logger.info('Segregating issues based on CONFIG')
        if issues:
            for issue in issues:
                print(issue)
                lastCommentDate = self.checklastcomment(issue, CONFIG.get("Comment Type"))
                if lastCommentDate != 0:
                    UTCdate = self.getdate(lastCommentDate)
                    lastComment = self.formatdate(lastCommentDate)
                else:
                    checkDate = issue.fields.created
                    UTCdate = self.getdate(checkDate)
                    lastComment = "Not Commented Yet"

                if issue.fields.status != "More Information":
                    dateCheck = self.comparedate(UTCdate, issue.fields.priority, CONFIG)
                    priority_msg = str(issue.fields.priority)
                else:
                    if CONFIG.get("More Information") != "None":
                        dateCheck = self.comparedate(UTCdate, "More Information", CONFIG)
                        priority_msg = "More Information"
                    else:
                        dateCheck = 0
                if issue.fields.assignee is None:
                    assignee = "Unassigned"
                else:
                    assignee = issue.fields.assignee.displayName

                if dateCheck == 1:
                    ticket = {
                        "ticket_number": issue.key,
                        "priority": priority_msg,
                        "assignee": assignee,
                        "status": issue.fields.status,
                        "lastcomment": str(lastComment)
                    }
                    self.tickets.append(ticket)
        else:
            logger.warning(msg='No Tickets found')

    @staticmethod
    def getexternalcomment(issueKey, commentId):
        """
            gets the type of comment whether it's Internal or Public
        :param issueKey: Issue ID (string)
        :param commentId: Comment ID (string)
        :return: Type of comment ('True/False')
        """
        url = f"{server}/rest/api/3/issue/{issueKey}/comment/{commentId}"
        auth = HTTPBasicAuth(email, api_token)
        headers = {
            "Accept": "application/json",
        }
        response = requests.get(url, headers=headers, auth=auth)
        data = json.loads(response.text)
        return data['jsdPublic']

    def checklastcomment(self, issue, commentConfig):
        """
            To get the last comments based on the comment Config
        :param issue: Issue ID (string)
        :param  commentConfig: Type of comment to check (string)
        :return: Date of the comment based on comment config or 0 (date object or 0)
        """
        global lastcommentdate
        comments = self.jira.comments(issue, expand="properties")
        if len(comments) != 0:
            comments.reverse()
            for comment in comments:
                if commentConfig == 'internal':
                    if not self.getexternalcomment(issue, str(comment)):
                        lastcommentdate = comment.created
                        break
                    else:
                        lastcommentdate = 0
                elif commentConfig == 'external':
                    if self.getexternalcomment(issue, str(comment)):
                        lastcommentdate = comment.created
                        break
                    else:
                        lastcommentdate = 0
                else:
                    lastcommentdate = comments[0].created
            if lastcommentdate is not None:
                return lastcommentdate
            else:
                return 0
        else:
            return 0

    @staticmethod
    def validateConfig(CONFIG):
        """
        To validate the CONFIG generated for the project
        :param CONFIG: CONFIG dictionary (dict)
        :return: 0 0r 1 based on the validity (0/1)
        """
        projectID = CONFIG.get("Project ID")
        P1 = CONFIG.get("1 - Critical")
        P2 = CONFIG.get("2 - Severe")
        P3 = CONFIG.get("3 - Moderate")
        P4 = CONFIG.get("4 - Minor")
        moreInfo = CONFIG.get("More Information")
        commentConfig = CONFIG.get("Comment Type")
        mailID = CONFIG.get("sender")
        if projectID == "None" or not projectID.isnumeric():
            logger.error("Invalid Project ID")
            return 0
        if mailID == "None":
            logger.error("Invalid Mail ID")
            return 0
        pattern = "([0-9]?[0-9])|None"
        if not re.match(pattern, P1) or P1 == 0:
            logger.error("Invalid days given for P1")
            return 0
        if not re.match(pattern, P2) or P2 == 0:
            logger.error("Invalid days given for P2")
            return 0
        if not re.match(pattern, P3) or P3 == 0:
            logger.error("Invalid days given for P3")
            return 0
        if not re.match(pattern, P4) or P4 == 0:
            logger.error("Invalid days given for P4")
            return 0
        if not re.match(pattern, moreInfo) or moreInfo == 0:
            logger.error("Invalid days given for Status More Information")
            return 0
        if commentConfig not in ["internal", "external", "all"]:
            logger.error("Invalid Comment Config given ")
            return 0
        return 1


    @staticmethod
    def comparedate(date, priority, CONFIG):
        """
        To Compare the last comment date with the current date
        :param date: date which needed to be compared (date object)
        :param priority: priority of the issue (Jira Object/String)
        :param CONFIG: CONFIG of the project (dict)
        :return: 1 or 0 based on the compared date results(0/1)
        """
        if CONFIG.get(str(priority)) != "None":
            days = int(CONFIG.get(str(priority)))
            checkDate = date + datetime.timedelta(days=days)
            today = datetime.datetime.utcnow().replace(microsecond=0)
            if checkDate.isoweekday() == 6:
                temp = (checkDate + datetime.timedelta(days=2)) - today
                if temp.days < 0:
                    return 1
                else:
                    return 0
            elif checkDate.isoweekday() == 7:
                temp = (checkDate + datetime.timedelta(days=1)) - today
                if temp.days < 0:
                    return 1
                else:
                    return 0
            else:
                temp = checkDate - today
                if temp.days < 0:
                    return 1
                else:
                    return 0
        else:
            return 0

    @staticmethod
    def sendmail(userID, tableContent):
        """
        To Send Emails automatically via OutLook
        :param userID: Sender ID for the mail(string)
        :param tableContent: List of Tickets need to be added to the MAIL(List)
        """
        try:
            logger.info('Creating Outlook Mail')
            outlook = client.Dispatch("Outlook.Application")
            message = outlook.CreateItem(0)
            message.To = userID
            attachment = message.Attachments.Add("logo.png")
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo_img")
            message.Subject = "Jira Reminder Tool "
            message.Body = message.HTMLBody = HTML_BODY(tableContent)
            message.Send()
            logger.info('Outlook Mail Send')
        except Exception:
            logger.error('Could not send mail')
            pass

    @staticmethod
    def getdate(string):
        """
        converts STRING object to a DATE object in UTC time zone
        :param string: date (string)
        :return: Date object
        """
        year = int(string[0:4])
        month = int(string[5:7])
        day = int(string[8:10])
        hour = int(string[11:13])
        minute = int(string[14:16])
        seconds = int(string[17:19])
        TZD = string[23:28]
        date = datetime.datetime(year=year, month=month, day=day, hour=hour, minute=minute, second=seconds)
        if TZD == "-0800":
            final = date + datetime.timedelta(hours=8)
        else:
            final = date + datetime.timedelta(hours=7)
        return final

    @staticmethod
    def gettable(tickets):
        """
        takes the tickets and converts it to HTML Table format
        :param tickets: List of tickets (list)
        :return: HTML Table version of the list of tickets (list)
        :return: HTML Table version of the list of tickets (list)
        """
        count = 0
        logger.info("Creating HTML Table for the mail")
        rows = ""
        for ticket in tickets:
            ticket_number = ticket.get("ticket_number")
            priority = ticket.get("priority")
            assignee = ticket.get("assignee")
            status = ticket.get("status")
            lastcomment = ticket.get("lastcomment")
            href = f"{server}/browse/{ticket_number}"
            temp = f"<tr><td><a href = \"{href}\">{ticket_number}</a></td><td>{priority}</td><td>{assignee}</td>" \
                   f"<td>{status}</td><td>{lastcomment}</td></tr>"
            rows = rows + temp
            count = count + 1
        print(count)
        logger.info(f'Total issues needed attention = {count}')
        logger.info('Created HTML table for the mail')
        return rows

    @staticmethod
    def formatdate(date):
        """
            To Change the Date Format for our convenience
            :param date: date(string)
        """
        year = int(date[0:4])
        month = int(date[5:7])
        day = int(date[8:10])
        formatedDate = f"{day}-{month}-{year}"
        return formatedDate

    @staticmethod
    def getnode(parentNode, tagName):
        """
        returns a list of nodes with the given tagname under the parent node, if it's none returns none.
        :param parentNode: node object
        :param tagName: child tag name(string)
        :return: child node object
        """
        try:
            if tagName == 'logger':
                if parentNode.getElementsByTagName(tagName)[0].firstChild is None:
                    logger.warning(f"Logger level set to ERROR")
                    return "None"
                else:
                    return parentNode.getElementsByTagName(tagName)[0].firstChild.nodeValue
            else:
                if parentNode.getElementsByTagName(tagName)[0].firstChild is None:
                    logger.warning(f"{tagName} is set to NONE")
                    return "None"
                else:
                    return parentNode.getElementsByTagName(tagName)[0].firstChild.nodeValue
        except Exception as e:
            if str(e) == "list index out of range":
                error = f"Tag {tagName} isn't available. Exception encountered : {e}"
                logger.error(error)
            else:
                logger.error(f"Exception encountered : {e}")
            return "invalid config"

    @staticmethod
    def getlogger(log):
        """
        Sets logger level
        :param log: log level need to be set (string)
        """
        if log in ['DEBUG', 'INFO', 'WARNING', 'CRITICAL', 'ERROR']:
            if log == 'DEBUG':
                logger.setLevel(logging.DEBUG)
                handler.setLevel(logging.DEBUG)
            if log == 'INFO':
                logger.setLevel(logging.INFO)
                handler.setLevel(logging.INFO)
            if log == 'WARNING':
                logger.setLevel(logging.WARNING)
                handler.setLevel(logging.WARNING)
            if log == 'CRITICAL':
                logger.setLevel(logging.CRITICAL)
                handler.setLevel(logging.CRITICAL)
            if log == 'ERROR':
                logger.setLevel(logging.ERROR)
                handler.setLevel(logging.ERROR)
            logger.addHandler(handler)
        else:
            logger.warning("Logger level is not set properly")



if __name__ == "__main__":
    jiraTool = JiraTool()
    pass
