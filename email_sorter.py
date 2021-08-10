import pandas as pd


input_file = "vba_programmer.xlsx"
in_sheet = pd.read_excel("vba_programmer.xlsx", sheet_name="Input")

final_output = pd.DataFrame(columns=['CustomerID', 'EmailSummary'])


def sort_emails(emails, email_summary, how_many_left, limit_hit):

    output = email_summary

    temp_count = 0
    temp_len = 0

    while (temp_count < len(emails)) and (temp_len < (64 - 7 - len(str(how_many_left)))):

        temp_out = output

        if limit_hit:
            return output, how_many_left, limit_hit

        try:
            if how_many_left > 1:
                temp_out = temp_out + emails[temp_count] + ';'
            else:
                temp_out = temp_out + emails[temp_count]
        except IndexError:
            pass

        if len(output) <= (64 - 7 - len(str(how_many_left))):

            if len(temp_out) > (64 - 7 - len(str(how_many_left))):
                output = f'{output}+ {how_many_left} more'
                limit_hit = True
            else:
                output = temp_out
                how_many_left = how_many_left - 1

        temp_len = len(output)
        temp_count = temp_count + 1

    return output, how_many_left, limit_hit


unique_customers = sorted(in_sheet['CustomerID'].unique())

for customer in unique_customers:
    temp = in_sheet[in_sheet['CustomerID'] == customer]

    primary_yes = []
    primary_no = []

    for index, row in temp.iterrows():

        if row['IsPrimary'] == 'yes':
            primary_yes.append(row['Email'])

        else:
            primary_no.append(row['Email'])

    customer_output = ''
    primary_yes = sorted(primary_yes)
    primary_no = sorted(primary_no)
    total_customers = len(primary_yes) + len(primary_no)
    limit = False

    customer_output, total_customers, limit = sort_emails(primary_yes, customer_output, total_customers, limit)
    customer_output, total_customers, limit = sort_emails(primary_no, customer_output, total_customers, limit)

    final_output = final_output.append({'CustomerID': customer, 'EmailSummary': customer_output}, ignore_index=True)


writer = pd.ExcelWriter('andrew_output_file.xlsx', engine='xlsxwriter')
final_output.to_excel(writer, sheet_name='my_output')
writer.save()
