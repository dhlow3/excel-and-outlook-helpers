# -*- coding: utf-8 -*-
"""Run SQL query, output results to CSV, attach to Email, and send."""
import os
import win32com.client
import sql_stuff


def email_data_feed(sqlpath, email_temp, csv_dst):
    """Attach a data file created from SQL query to an email template and send.

    Parameters
    ----------
    sqlpath: str
        Path to the .sql file to be executed.
    email_temp: str
        Path to an Outlook .msg email.
    csv_dst: str
        Path to the .csv file that will be created from query results.
    """
    assert os.path.exists(sqlpath)
    assert os.path.exists(email_temp)
    assert isinstance(sqlpath, str)
    assert isinstance(email_temp, str)
    assert isinstance(csv_dst, str)

    # Connect to SQL engine
    sql = sql_stuff.SQLConnection()

    # Run query and return results as Pandas DataFrame
    orders_df = sql.execute_sql(query_input=sqlpath, return_results=True)

    # Output DataFrame as .CSV file
    orders_df.to_csv(csv_dst, header=True, index=False)

    # Create instance of Outlook
    obj = win32com.client.Dispatch('Outlook.Application')

    # Define template email file
    mail = obj.CreateItemFromTemplate(TemplatePath=email_temp)

    # Attach file
    mail.Attachments.Add(csv_dst)

    # Send Email
    mail.Send()


def main():
    # Example use
    MAIN_DIR = os.path.dirname(__file__)
    SQL_FILE = 'test.sql'  # .sql filename
    EMAIL_TEMP = 'test.msg'  # .msg filename
    OUTPUT_FILENAME = 'test.csv'  # .csv filename

    # Define file relative locations
    sqlpath = os.path.join(MAIN_DIR, SQL_FILE)
    email_temp = os.path.join(MAIN_DIR, EMAIL_TEMP)
    csv_dst = os.path.join(MAIN_DIR, OUTPUT_FILENAME)

    email_data_feed(sqlpath, email_temp, csv_dst)


if __name__ == '__main__':
    main()
