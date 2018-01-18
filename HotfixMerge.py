import const
import gspread
from oauth2client.service_account import ServiceAccountCredentials as SAC
import P4Summary
from jira import JIRA

def main():
    print('Update Spreadsheet for merged Hotfix...')
    # import time
    # start_time = time.time()

    scope = [const.google_sheet_url]
    credentials = SAC.from_json_keyfile_name(const.google_json_file, scope)
    gc = gspread.authorize(credentials)
    worksheet_60 = gc.open_by_key(const.google_sheet_key).worksheet('[EN] 6.0 SP3 Patch 3')
    worksheet_70 = gc.open_by_key(const.google_sheet_key).worksheet('[EN] 7.0')

    next_data = P4Summary.exist_hotfix(worksheet_60)
    row_count = next_data[0]

    # caseID = 'SEG-18510'
    # print(get_link_from_jira(caseID))

    # update
    for cursor_row_count in range(2, row_count):
        readme_hotfix_number = P4Summary.get_hotfix_number(worksheet_60.cell(cursor_row_count, 2).value)
        if readme_hotfix_number != '' and int(readme_hotfix_number) > 3686:    # the last hotfix merged intto TMCM 7.0 GM from TMCM 6.0
            cases = (worksheet_60.cell(cursor_row_count, 3).value).split('\n')
            tmp_merge_link = worksheet_60.cell(cursor_row_count, 12).value
            if len(cases) != tmp_merge_link.count('HFB'):
                for case in cases:
                    linked_case = get_link_from_jira(case).strip('\n')
                    if len(linked_case) > 0 and (linked_case + ' (HFB') not in tmp_merge_link:
                        merged_hotfix =  get_hotfix_from_spreadsheet(linked_case, worksheet_70)
                        if case not in tmp_merge_link:
                            tmp_merge_link += '[' + case + ']' + '\n' + linked_case + ' (' + merged_hotfix + ')' + '\n'
                            worksheet_60.update_cell(cursor_row_count, 12, tmp_merge_link.strip('\n'))
                        elif (linked_case + ' ()') in tmp_merge_link and merged_hotfix.strip('\n') != '' :
                            tmp_merge_link = str(tmp_merge_link).replace(linked_case + ' ()', linked_case + ' (' + merged_hotfix + ')' )
                            worksheet_60.update_cell(cursor_row_count, 12, tmp_merge_link.strip('\n'))


    # print("time elapsed: {:.2f}s".format(time.time() - start_time))


def get_link_from_jira(ticket_number):
    jira = JIRA(server=const.jira_server, basic_auth=(const.jira_account, const.jira_password))
    issue = jira.issue(ticket_number)
    link_list = ''
    for link in issue.fields.issuelinks:
        if hasattr(link, "outwardIssue"):
            outwardIssue = jira.issue(str(link.outwardIssue))
            if 'Control Manager' in outwardIssue.fields.customfield_11103 and '7' in str(outwardIssue.fields.customfield_11104):
              link_list += outwardIssue.key + '\n'
        if hasattr(link, "inwardIssue"):
            inwardIssue = jira.issue(str(link.inwardIssue))
            if 'Control Manager' in inwardIssue.fields.customfield_11103 and '7' in str(inwardIssue.fields.customfield_11104):
              link_list += inwardIssue.key + '\n'

    return(link_list)


def get_hotfix_from_spreadsheet(ticket_number, worksheet):
    next_data = P4Summary.exist_hotfix(worksheet)
    row_count = next_data[0]
    result = ''
    for cursor_row_count in range(2, row_count):
        if ticket_number in worksheet.cell(cursor_row_count, 3).value:
            result += worksheet.cell(cursor_row_count, 2).value

    return result

    
if __name__ == "__main__":
    main()
