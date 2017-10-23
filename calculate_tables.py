import pandas as pd
from utils import * 
import sys
import settings

def main():
    '''
    Return - Ouput excel file - daily production and Average Cost of Production 
                Replacement (12 mo Trailing in Qtrly) tables
                Exit production numbers for this and last years.

    Input - csv file - EXAMCSV_1.csv. Please use Generic Mode on Home Table when you save.
                    Otherwise, it will not work.(I did not have much time to clean file for other modes)
            Output excel file - Please use excel(xlsx) file. Otherwise, it cannot create output file.

            Example : python calculate_tables.py EXAMCSV_1.csv OutputFile.xlsx
    '''

    try:
        fName=sys.argv[1]
        outFile=sys.argv[2]
    except:
        print('You failed to provide all inputs on the command line!')
        print('Give input csv. Output file must be xlsx file')
        print('Example: python cal.py EXAMCSV_1.csv OurFile.xlsx')
        sys.exit(1)  # abort

    try:
        days_dic,df_const,df_capex = reading_file(fName)
    except : 
        print('Check your input csv file! Please use Generic Mode on Home Table when you save csv file.')
        sys.exit(1)  # abort

    settings.days_dic=days_dic
    year_list= sorted(set(df_capex.Year.values))
    LastYear=year_list[-1]
    
    for year in year_list:
        
        yearInt=int(year)

        try :
            settings.const_dic=df_const[df_const.Year==year][['Type2','Value']].set_index('Type2').to_dict()['Value']
        except :
            print('Some constant numbers are missing!')
            print('check year ',year)
            sys.exit(1)
        

        df_cap=creating_capital_spending(df_capex, year)

        # prod
        #df_prod=create_production_forecast(df_cap, last_years_closing, const_dic)
        df_prod=create_production_forecast(df_cap)
        # boe summary
        df_summary=create_boe_summary(df_prod)
        df_oil=create_oil_summary(df_summary)

        df_gas=create_gas_summary(df_summary)

        # Daily Production
        df_daily=create_daily_production(df_gas,df_oil)

        # Annual Production
        df_annual=create_annual_production(df_daily)

        # cannot produce the table for 2013 because it does not have enough info
        if year !='2013':
            df_ave_prod=create_avecost_production(df_prod, df_prod_last, \
                                           df_cap, df_cap_last,df_daily) 

        
        if year == LastYear:
            # writing files
            exit_prod=[df_prod_last.loc['Closing for period','Total'],df_prod.loc['Closing for period','Total']]
            if year == '2014':
                create_excelfile( yearInt, outFile, df_daily_last,
                                df_ave_prod,df_daily,df_ave_prod,exit_prod)
            else:
                create_excelfile( yearInt, outFile, df_daily_last,
                                df_ave_prod_last,df_daily,df_ave_prod,exit_prod)

        if yearInt > 2013:
            df_ave_prod_last=df_ave_prod 
                


        # needed tables for next years calculation
        if yearInt > 2013:
            df_ave_prod_last=df_ave_prod 

        df_prod_last=df_prod
        df_cap_last=df_cap
        df_daily_last=df_daily
        # others
        df_oil_last=df_oil
        df_gas_last=df_gas
        df_summary_last=df_summary
        df_annual_last=df_annual

    
        settings.last_years_closing['Year']=str(yearInt+1)
        settings.last_years_closing['CapEff']=settings.const_dic['CapEff']
        settings.last_years_closing['TieIn']=settings.const_dic['TieIn']

        settings.last_years_closing['Closing for period']=df_prod.loc['Closing for period','Total']
        settings.last_years_closing['Adds - previous carryover']=df_prod.loc['Adds - previous carryover','Q4']
        settings.last_years_closing['Capex Drilling']=df_prod.loc['Capex Drilling','Q4']

        # I have to code it because data for 2013 are modified.

        settings.last_years_closing['Declines on previous year carryover'] = 0 # it was -200 for 2013. But, it is set to zero for 2014 and after.

        if year=='2013':
            settings.last_years_closing['NGLs (Bbls/d)'] = 336.7  # Because some numbers for 2013 are tweaked, I have to use the numbers given in the example file.
            settings.last_years_closing['Natural Gas (MMcf/d)'] = 19.7
            settings.last_years_closing['Total (Boe/d 6:1)'] = 6872.0
        else :
            settings.last_years_closing['NGLs (Bbls/d)'] = df_daily_last.loc['NGLs (Bbls/d)','Q4']  # After 2014, it will use the number given in the daily prod table. 
            settings.last_years_closing['Natural Gas (MMcf/d)'] = df_daily_last.loc['Natural Gas (MMcf/d)','Q4']
            settings.last_years_closing['Total (Boe/d 6:1)'] = df_daily_last.loc['Total (Boe/d  6:1)','Total']


if __name__ == "__main__":

    main()

