"""
Created on Thu Dec 30 18:51:03 2021

@author: Rodney Nobles
@WebAPI: Rodney Nobles
"""
from datetime import datetime
import pandas as pd
import time
import anvil.server
import anvil.media
import smtplib
import ssl
from datetime import datetime

anvil.server.connect("546EZXJHJUXEJBQAFPVPOCSF-SIBOH5Y3CE6BIS3D")
# In a script on your own machine (or anywhere)


@anvil.server.callable
def get_output(spdr_file_name, mspr_file_name, spdr_file, mspr_file, territory, ownership, address):
    '''Generate output.csv from spdr_file, mspr_file, territory, ownership
    '''

    start_time = time.time()
    now = datetime.now()
    current_time = now.strftime("%d/%m/%Y %H:%M:%S")
    print(f"Timestamp: {current_time}")
    # send email
    print(f"User: {address}")
    print(ownership)
    print(f"mspr file name: {mspr_file_name}")
    print(f"spdr file name: {spdr_file_name}")

    ownership_dict = create_ownership_dict(spdr_file, spdr_file_name)
    # Create output df
    output_df = create_output_df()

    # Open mspr file and iterate through stores
    if "csv" in mspr_file_name:
        with anvil.media.TempFile(mspr_file) as mspr_f:
            mspr_df = pd.read_csv(mspr_f, dtype={9: str})
            output_df = iterate_mspr(
                output_df, mspr_df, territory, ownership_dict, ownership)
    if "xlsx" in mspr_file_name:
        with anvil.media.TempFile(mspr_file) as mspr_f:
            mspr_df = pd.read_excel(mspr_f, dtype={9: str})
            output_df = iterate_mspr(
                output_df, mspr_df, territory, ownership_dict, ownership)

    # output Metro Store Perfomrance Report"
    output_filename = 'Output_Data.xlsx'
    output_df.to_excel(output_filename,  index=False)

    # send to user
    output_file = anvil.media.from_file(output_filename)
    print("--- %s seconds ---" % (time.time() - start_time))
    return output_file


def iterate_mspr(output_df, mspr_df, territory, ownership_dict, ownership):
    for index, row in mspr_df.iterrows():
        # If Billing state/province in territory?
        if check_territory(row, territory):
            # Door code can be found in ownership dctionary?
            if row[5] in ownership_dict[ownership]:
                # Final output
                output_df = append_to_output_file(output_df, ownership, row)
    return output_df


def add_values_in_dict(sample_dict, key, list_of_values):
    ''' Append multiple values to a key in 
        the given dictionary '''
    if key not in sample_dict:
        sample_dict[key] = list()
    sample_dict[key].extend(list_of_values)
    return sample_dict


def create_ownership_dict(spdr_file, spdr_file_name):
    ''' Read SPDR.xslx as python df &
        Create spdr dictionary with ownership key and list of door codes as value
    '''
    #D(ownership), Q(doorcode)
    if "xlsx" in spdr_file_name:
        with anvil.media.TempFile(spdr_file) as spdr_f:
            spdr_df = pd.read_excel(spdr_f, usecols=[3, 16])
            ownership_dict = {}
            for index, sp_row in spdr_df.iterrows():
                # sp_row[0](ownership), sp_row[1](doorcode)
                ownership_dict = add_values_in_dict(
                    ownership_dict, sp_row[0], [sp_row[1]])
    elif "csv" in spdr_file_name:
        with anvil.media.TempFile(spdr_file) as spdr_f:
            spdr_df = pd.read_csv(spdr_f, usecols=[3, 16])
            ownership_dict = {}
            for index, sp_row in spdr_df.iterrows():
                # sp_row[0](ownership), sp_row[1](doorcode)
                ownership_dict = add_values_in_dict(
                    ownership_dict, sp_row[0], [sp_row[1]])
    return ownership_dict


def create_output_df():
    ''' Create output DataFrame and Headers
    '''
    output_df = pd.DataFrame(columns=['Ownership',
                                      'Email',
                                      'Door Code',
                                      'Address',
                                      'Billing City',
                                      'Billing State/Province',
                                      'Billing Zip',
                                      'Apps Submitted CM',
                                      'Apps Approved CM',
                                      'Transactions CM',
                                      'Take Rate CM',
                                      'Originations CM',
                                      'Average Lease Total',
                                      'Apps Submitted PM',
                                      'Apps Approved PM',
                                      'Transaction PM',
                                      'Originations PM',
                                      'Missed Opps PM',
                                      'PY Total Transactions',
                                      'PY Total Originations',
                                      'PY Total Apps',
                                      'PY Total Approved Apps',
                                      'PY Total Missed Opps',
                                      ])
    return output_df


def check_territory(row, territory):
    ''' Check if Billing state/province in territory?
    '''
    state_abbreviation = str(row[11])
    state_abbreviation = state_abbreviation[0:2]
    if state_abbreviation in territory:
        if state_abbreviation == "PA":
            if row[2] == "Pittsburgh":
                in_territory = True
            else:
                in_territory = False
        else:
            in_territory = True
    else:
        in_territory = False
    return in_territory


def average_lease_total(row):
    '''Average Lease Total
    '''
    if ((row[20]) == ""):
        average_lease_total = "N/A"
    elif (float(row[20]) == 0):
        average_lease_total = 0
    else:
        average_lease_total = float(row[23])/float(row[20])
    return average_lease_total


def append_to_output_file(output_df, ownership, row):
    '''Append row to output
    '''
    output_df = output_df.append({'Ownership': ownership,
                                  'Email': row[7],
                                  'Door Code': row[5],
                                  'Address': row[8],
                                  'Billing City': row[10],
                                  'Billing State/Province': row[11],
                                  'Billing Zip': row[12],
                                  'Apps Submitted CM': row[17],
                                  'Apps Approved CM': row[18],
                                  'Transactions CM': row[20],
                                  'Take Rate CM': row[22],
                                  'Originations CM': row[23],
                                  'Average Lease Total': average_lease_total(row),
                                  'Apps Submitted PM': row[25],
                                  'Apps Approved PM': row[26],
                                  'Transaction PM': row[28],
                                  'Originations PM': row[30],
                                  'Missed Opps PM': row[32],
                                  'PY Total Transactions': row[47],
                                  'PY Total Originations': row[48],
                                  'PY Total Apps': row[49],
                                  'PY Total Approved Apps': row[50],
                                  'PY Total Missed Opps': row[51]},
                                 ignore_index=True)
    return output_df


anvil.server.wait_forever()
