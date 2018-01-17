# -*- coding: utf-8 -*-
"""Refresh connections and pivot caches in Excel workbook."""
import os
import sys

from win32com.client import DispatchEx, gencache
from time import sleep


def refresh_connection(excel, workbook, connection_type):
    """Refresh a connection in Excel workbook.

    Parameters
    ----------
    excel: win32com.gen_py.xxx._Application
        An instance of Excel.
    workbook: win32com.gen_py.xxx.Workbook
        Excel workbook object.
    connection_type: xlConnectionType
        See xlconnectiontype-enumeration-excel in Microsoft's documentation.
    """
    # Iterate workbook connections and refresh if connection type matches
    for connection in workbook.Connections:
        if connection.Type == connection_type:

            # OLEDB connections
            if connection.Type == 1:
                connection.OLEDBConnection.Refresh()
                while True:
                    # Ensure refresh is complete before continuing
                    if connection.OLEDBConnection.Refreshing:
                        sleep(1)
                    else:
                        break
                print('-{} connection refreshed (type {})'.format(
                      connection.Name, connection_type))

            # Other connection types
            else:
                connection.Refresh()
                print('-{} connection refreshed (type {})'.format(
                      connection.Name, connection_type))

    # Run all pending queries to OLEDB and OLAP data sources
    excel.CalculateUntilAsyncQueriesDone()


def refresh_pivot_caches(excel, workbook):
    """Refresh Pivot caches.

    Parameters
    ----------
    excel: win32com.gen_py.xxx._Application
        An instance of Excel.
    workbook: win32com.gen_py.xxx.Workbook
        Excel workbook object.
    """
    pivot_caches = workbook.PivotCaches
    for cache in pivot_caches():
        if cache.SourceType != 2:  # external sources
            cache.Refresh()
            print('-Pivot cache (type {}) refreshed'.format(cache.SourceType))

    # Run all pending queries to OLEDB and OLAP data sources
    excel.CalculateUntilAsyncQueriesDone()


def main():
    FILE_PATH = sys.argv[1]
    assert os.path.exists(FILE_PATH)

    # Use Excel 2016 library; change/remove if not using 2016
    gencache.EnsureModule(
        '{00020813-0000-0000-C000-000000000046}', 0, 1, 9
        )

    # Create instance of Excel
    xl = DispatchEx('Excel.Application')

    # Set application properties
    xl.Visible = 0
    xl.DisplayAlerts = False

    # Create workbook object
    workbook = xl.Workbooks.Open(FILE_PATH)

    # Loop to see if workbook is ready to proceed
    workbook_ready = False
    tries = 0

    while not workbook_ready:
        try:
            workbook.Activate()
        except Exception:
            sleep(1)
            tries += 1
            if tries == 10:
                break
        else:
            workbook_ready = True

    """For silent refresh operations, use the FastCombine property in
    conjunction with the Application.DisplayAlerts property, set to False"""
    workbook.Queries.FastCombine = True

    # Common connection types; add more if needed
    connection_types = {'1_ODBC': 2,
                        '2_OLEDB': 1,
                        '3_MODEL': 7,
                        '4_WORKSHEET': 8}

    # Refresh connections in a specific order
    for connection_type in sorted(connection_types.keys()):
        refresh_connection(excel=xl,
                           workbook=workbook,
                           connection_type=connection_types[connection_type])

    # Disable events while caches refresh
    xl.EnableEvents = False
    refresh_pivot_caches(excel=xl, workbook=workbook)
    xl.EnableEvents = True

    # Calculate any necessary calculations in the workbook
    xl.Calculate()

    # Close workbook object
    workbook.Close(True)
    # Close Excel instance
    xl.Quit()
    # Remove Excel instance
    del xl


if __name__ == '__main__':
    main()
