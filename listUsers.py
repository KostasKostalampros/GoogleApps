import os

import csv
import httplib2
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

# Script universal variables
SCOPES = 'https://www.googleapis.com/auth/admin.directory.user'
CLIENT_SECRET_FILE = 'CLIENT_SECRET_FILEPATH'
APPLICATION_NAME = 'Google Directory List Users'


try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None


def get_credentials():
    """Gets valid user credentials from storage.

     If nothing has been stored, or if the stored credentials are invalid,
     the OAuth2 flow is completed to obtain the new credentials.

     Returns:
         Credentials, the obtained credential.
     """

    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'P:\\service_acccount_keys\\admin-directory_v1-python-quickstart.json')
    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
    return credentials


def parse_results(results):
    """Parse results and populate one list per column from the Google Apps API

    Args:
        results: The results returned from the API call execution.

    Returns:
        Lists of the parsed column results.
    """

    # Initiate all column lists
    list_name = []
    list_suspended = []
    list_suspendedReason = []
    list_primary_email = []
    list_all_emails = []
    list_isAdmin = []
    list_isDelegatedAdmin = []
    list_lastLoginTime = []
    list_creationTime = []
    list_job_title = []
    list_job_department = []
    list_job_location = []

    # For each user in the results get the values and append in the respective lists defined above
    for user in results.get('users'):
        user_full_name = user.get('name').get('fullName').encode("utf-8")
        user_suspended = user.get('suspended')
        user_suspended_reason = user.get('suspensionReason')
        user_primary_email = user.get('primaryEmail')

        user_all_emails = ""
        for email in user.get('emails'):
            user_all_emails += email.get('address') + ";"

        user_isAdmin = user.get('isAdmin')
        user_isDelegatedAdmin = user.get('isDelegatedAdmin')
        user_lastLoginTime = user.get('lastLoginTime')
        user_creationTime = user.get('creationTime')
        try:
            user_job_title = user.get('organizations')[0].get('title')
        except (AttributeError, TypeError):
            user_job_title = ""
        try:
            user_job_department = user.get('organizations')[0].get('department')
        except (AttributeError, TypeError):
            user_job_department = ""
        try:
            user_job_location = user.get('organizations')[0].get('location')
        except (AttributeError, TypeError):
            user_job_location = ""

        list_name.append(user_full_name)
        list_suspended.append(user_suspended)
        list_suspendedReason.append(user_suspended_reason)
        list_primary_email.append(user_primary_email)
        list_all_emails.append(user_all_emails)
        list_isAdmin.append(user_isAdmin)
        list_isDelegatedAdmin.append(user_isDelegatedAdmin)
        list_lastLoginTime.append(user_lastLoginTime)
        list_creationTime.append(user_creationTime)
        list_job_title.append(user_job_title)
        list_job_department.append(user_job_department)
        list_job_location.append(user_job_location)

    return list_name, list_suspended, list_suspendedReason, list_primary_email, list_all_emails, list_isAdmin, \
           list_isDelegatedAdmin, list_lastLoginTime,list_creationTime, list_job_title, list_job_department, \
           list_job_location


def get_user_data_list(service):
    """ Create calls to the Google Apps API and populate the column lists

    Args:
        service: The service that is connected to the specified API.

    Returns:
        List of lists with the requested user data.
    """

    # Initiate all column lists with headers
    list_name = ['Name']
    list_suspended = ['Suspended']
    list_suspendedReason = ['SuspendedReason']
    list_primary_email = ['PrimaryEmailAddress']
    list_all_emails = ['EmailAddress']
    list_isAdmin = ['IsAdmin']
    list_isDelegatedAdmin = ['IsDelegatedAdmin']
    list_lastLoginTime = ['LastLoginTime']
    list_creationTime = ['CreationTime']
    list_job_title = ['Title']
    list_job_department = ['Department']
    list_job_location = ['Location']

    # Make the API call
    results = service.users().list(customer='my_customer', maxResults=500,
                                   orderBy='familyName').execute()

    # Send the API results for parsing and store output in temporary lists
    list_name_temp, list_suspended_temp, list_suspendedReason_temp, list_primary_email_temp, list_all_emails_temp, list_isAdmin_temp, \
    list_isDelegatedAdmin_temp, list_lastLoginTime_temp, list_creationTime_temp, list_job_title_temp, list_job_department_temp, \
    list_job_location_temp = parse_results(results)

    # Add the temporary lists to the respective main lists which in the end populate the final data list
    list_name.extend(list_name_temp)
    list_suspended.extend(list_suspended_temp)
    list_suspendedReason.extend(list_suspendedReason_temp)
    list_primary_email.extend(list_primary_email_temp)
    list_all_emails.extend(list_all_emails_temp)
    list_isAdmin.extend(list_isAdmin_temp)
    list_isDelegatedAdmin.extend(list_isDelegatedAdmin_temp)
    list_lastLoginTime.extend(list_lastLoginTime_temp)
    list_creationTime.extend(list_creationTime_temp)
    list_job_title.extend(list_job_title_temp)
    list_job_department.extend(list_job_department_temp)
    list_job_location.extend(list_job_location_temp)

    # Check if there are more than 500 results in order to get the next page token
    if results.get('nextPageToken'):
        nextPageToken = results.get('nextPageToken')
    else:
        nextPageToken = ''

    # While there are more results identified from the next page tokens parse them and then add them to their respective lists
    while True:
        results = service.users().list(customer='my_customer', maxResults=500,
                                       orderBy='familyName', pageToken=nextPageToken).execute()

        list_name_temp, list_suspended_temp, list_suspendedReason_temp, list_primary_email_temp, list_all_emails_temp, list_isAdmin_temp, \
        list_isDelegatedAdmin_temp, list_lastLoginTime_temp, list_creationTime_temp, list_job_title_temp, list_job_department_temp, \
        list_job_location_temp = parse_results(results)

        list_name.extend(list_name_temp)
        list_suspended.extend(list_suspended_temp)
        list_suspendedReason.extend(list_suspendedReason_temp)
        list_primary_email.extend(list_primary_email_temp)
        list_all_emails.extend(list_all_emails_temp)
        list_isAdmin.extend(list_isAdmin_temp)
        list_isDelegatedAdmin.extend(list_isDelegatedAdmin_temp)
        list_lastLoginTime.extend(list_lastLoginTime_temp)
        list_creationTime.extend(list_creationTime_temp)
        list_job_title.extend(list_job_title_temp)
        list_job_department.extend(list_job_department_temp)
        list_job_location.extend(list_job_location_temp)

        if results.get('nextPageToken'):
            nextPageToken = results.get('nextPageToken')
        else:
            break

    # Generate final user data list of lists containing all the individual lists generated above
    final_user_data_list = [list_name, list_suspended, list_suspendedReason, list_primary_email, list_all_emails, list_isAdmin,
                            list_isDelegatedAdmin, list_lastLoginTime, list_creationTime, list_job_title, list_job_department,
                            list_job_location]

    return final_user_data_list


def print_table_to_csv(data_list, filename):
    """Print input list of lists in a csv format in the same folder as this scripts.

    Args:
        data_list: A list of lists populated with data.
        filename: A file name in string format used for the output file.

    Returns:
        None
    """
    with open(filename, "w", newline="") as file:
        writer = csv.writer(file, delimiter=",")
        for iter in range(len(data_list[0])):
            writer.writerow([(x[iter]) for x in data_list])


def main():

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('admin', 'directory_v1', http=http)

    output_csv_filename = "GoogleAppsUserDataList.csv"

    final_user_data_list = get_user_data_list(service)
    print_table_to_csv(final_user_data_list, output_csv_filename)


if __name__ == '__main__':
    main()