import pandas as pd
import os
import glob

# get all housing files for each survey and every metro area
path = os.getcwd()
all_surveys = glob.glob(os.path.join('FILEPATH FOR FOLDER WITH ALL SURVEYS', "*.xlsx"))

# list of desired regions
metros = ['Philadelphia_Metro_Area', 'San.Francisco_Metro_Area', 'Riverside_Metro_Area', 'Los.Angeles_Metro_Area',
          'Chicago_Metro_Area', 'New.York_Metro_Area', 'Houston_Metro_Area', 'NC', 'US']

# create dict of all metros' aggregated surveys
metro_surveys = {}

# loop through surveys from 15 weeks
for week in all_surveys:
    # loop through metros in our markets
    for metro in metros:
        # read in sheet of data for each metro in a given week
        current_metro_week = pd.read_excel(week, sheet_name=metro, names=['Characteristic', 'Total',
                                                                          'Reduced_Almost_Every_Month',
                                                                          'Reduced_Sometimes',
                                                                          'Reduced_One_Or_Two_Months',
                                                                          'Reduced_Never',  'Reduced_Did_Not_Report',
                                                                          'Unhealthy_Almost_Every_Month',
                                                                          'Unhealthy_Sometimes',
                                                                          'Unhealthy_One_Or_Two_Months',
                                                                          'Unhealthy_Never',
                                                                          'Unhealthy_Did_Not_Report',
                                                                          'Unable_Almost_Every_Month',
                                                                          'Unable_Sometimes',
                                                                          'Unable_One_Or_Two_Months', 'Unable_Never',
                                                                          'Unable_Did_Not_Report'])
        # drop top rows with extraneous info
        current_metro_week = current_metro_week.drop([0, 1, 2, 3, 4, 5])
        # loop through survey questions
        for field in current_metro_week.columns[1:]:
                # loop through respondent characteristics
                for n in range(7, 124):
                    # check if metro is already in overarching dict
                    if metro in metro_surveys:
                        # check if existing cell values are integers
                        if type(metro_surveys[metro][field][n]) is int:
                            # check if new cell values are integers
                            if type(current_metro_week[field][n]) is int:
                                # add numbers from current week to the metro's total for each field
                                metro_surveys[metro][field][n] += current_metro_week[field][n]
                        else:
                            # set value for new field when existing value is '-'/null
                            metro_surveys[metro][field][n] = current_metro_week[field][n]
                    else:
                        # add new metro to dict
                        metro_surveys[metro] = current_metro_week

# loop through each metro master sheet in dict
for name, df in metro_surveys.items():
    # clean up null cells
    df.fillna(0, inplace=True)
    df.replace('-', 0, inplace=True)
    # create column for percent of respondents who reduced necessary expenses to pay an energy bill
    df['Pct_Reduced'] = (df['Reduced_Almost_Every_Month'].apply(int) + df['Reduced_Sometimes'].apply(int) +
                         df['Reduced_One_Or_Two_Months'].apply(int)) / (df['Total'] -
                                                                        df['Reduced_Did_Not_Report']).apply(int) * 100
    # create column for percent of respondents who kept their home at an unhealthy temperature to pay an energy bill
    df['Pct_Unhealthy'] = (df['Unhealthy_Almost_Every_Month'].apply(int) + df['Unhealthy_Sometimes'].apply(int) +
                           df['Unhealthy_One_Or_Two_Months'].apply(int)) / \
                          (df['Total'].apply(int) - df['Unhealthy_Did_Not_Report'].apply(int)) * 100
    # create column for percent of respondents who were unable to pay an energy bill
    df['Pct_Unable'] = (df['Unable_Almost_Every_Month'].apply(int) + df['Unable_Sometimes'].apply(int) +
                        df['Unable_One_Or_Two_Months'].apply(int)) / (df['Total'].apply(int) -
                                                                      df['Unable_Did_Not_Report'].apply(int)) * 100
    # create column for likelihood ratio of respondents reducing necessary expenses (compared to overall rate)
    df['Reduced_Likelihood_Ratio'] = df['Pct_Reduced'] / df['Pct_Reduced'][6]
    # create column for likelihood ratio of respondents keeping home at unhealthy temperature (compared to overall rate)
    df['Unhealthy_Likelihood_Ratio'] = df['Pct_Unhealthy'] / df['Pct_Unhealthy'][6]
    # create column for likelihood ratio of respondents being unable to pay (compared to overall rate)
    df['Unable_Likelihood_Ratio'] = df['Pct_Unable'] / df['Pct_Unable'][6]
    # download each metro master sheet as a csv file
    df.to_csv('FILEPATH FOR LOCATION TO SAVE CSV' + name + '.csv')
