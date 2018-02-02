import const
import P4Summary
import chardet
import time

def main():
    # try:
    #     start_time = time.time()?
        global jira
        jira = const.connect_jira()

        worksheet_60 = const.connect_google_spreadsheet('[EN] 6.0 SP3 Patch 3')
        hotfix_list_60 = prepare_case_list(worksheet_60)
        worksheet_readme = const.connect_google_spreadsheet('[EN] 6.0 SP3 Patch 3 Readme')
        insert_hotfix(worksheet_readme, hotfix_list_60)
        update_readme(worksheet_readme)
    # finally:
    #     print("time elapsed: {:.2f}s".format(time.time() - start_time))

def prepare_case_list(worksheet_60):
    print('prepare_case_list...')
    row_count = 2
    hotfix_list = []
    while worksheet_60.cell(row_count, 1).value != '':
        row_data = worksheet_60.row_values(row_count)
        case_list = []
        for case in row_data[2].split('\n'):
            case_list.append(case)
        hotfix_list.append([P4Summary.get_hotfix_number(row_data[1]), case_list, row_data[9]])
        row_count += 1
    return hotfix_list

def insert_hotfix(worksheet_readme, hotfix_list_60):
    print('insert_hotfix...')
    next_data = exist_hotfix(worksheet_readme, 1)
    row_count = next_data[0]
    for hotfix in hotfix_list_60:
        if hotfix[0] > next_data[1]:
            for case in hotfix[1]:
                worksheet_readme.insert_row([hotfix[0], case, '', '', '', '', hotfix[2]], row_count)
                row_count += 1
    # print()

def update_readme(worksheet_readme):
    print('update_readme...')
    cursor_row_count = 2
    while worksheet_readme.cell(cursor_row_count, 2).value != '':
        case = (worksheet_readme.cell(cursor_row_count, 2).value)
        # print(case)
        if worksheet_readme.cell(cursor_row_count, 4).value == '' or worksheet_readme.cell(cursor_row_count, 5).value == '' or worksheet_readme.cell(cursor_row_count, 6).value == '':
            readme_case = get_readme_from_jira(case)
            if worksheet_readme.cell(cursor_row_count, 4).value == '' and readme_case['jp'] != '':
                worksheet_readme.update_cell(cursor_row_count, 4, readme_case['jp'])
            if worksheet_readme.cell(cursor_row_count, 5).value == '' and readme_case['sc'] != '':
                worksheet_readme.update_cell(cursor_row_count, 5, readme_case['sc'])
            if worksheet_readme.cell(cursor_row_count, 6).value == '' and readme_case['fr'] != '':
                worksheet_readme.update_cell(cursor_row_count, 6, readme_case['fr'])

        cursor_row_count += 1

def get_readme_from_jira(ticket_number):
    dict = {'case': ticket_number, 'jp': '', 'sc': '', 'fr': ''};
    if 'TT' not in ticket_number:
        issue = jira.issue(ticket_number)
        for attachment in issue.fields.attachment:
            if attachment.mimeType == 'text/plain' and 'readme' in attachment.filename.lower():
                if 'jp' in attachment.filename.lower():
                    dict['jp'] = get_readme_content(attachment)
                elif 'sc' in attachment.filename.lower():
                    dict['sc'] = get_readme_content(attachment)
                elif 'fr' in attachment.filename.lower():
                    dict['fr'] = get_readme_content(attachment)

    return dict

def get_readme_content(attachment):
    data = attachment.get()
    decode = chardet.detect(data)
    if len(data.decode(decode['encoding'])) > 50000:
        return 'Exceed 50,000 characters'
    else:
        return str(data.decode(decode['encoding']))

def exist_hotfix(worksheet, column):
    row_data = worksheet.col_values(column)
    row_data = [k for k in row_data if k != '']
    next_row = len(row_data) + 1

    if next_row > 2:
        row_data = sorted(row_data, reverse = True)
        return [next_row, row_data[1]]
    else:
        return [next_row, '']

if __name__ == "__main__":
    main()
