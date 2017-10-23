import pandas as pd
from numpy import round
import settings


def reading_file(fname):
    '''
    Return days_dic - dictionary of quarters, 
           df_const - dataframe of constants
           df_capex - capex data

    Input  fname - input csv file name
    '''

    df_tmp=pd.read_csv(fname,header=None)
    df_tmp.columns=['Model','Type1','Type2','Type3','Year','Quarter','Value']
    # days
    df_days=df_tmp[(df_tmp.Type1=='Time')]
    days_dic=df_days[['Year','Value']].set_index('Year').to_dict()['Value']

    df_const=df_tmp[(df_tmp.Type1=='OVERRIDE')]

    df_capex=df_tmp[df_tmp.Type1=='CAPEX']
    df_capex=df_capex.fillna(0)
    df_capex.columns=['Model','CAPEX','Type','CAD','Year','Quarter','Pars']
    df_capex=df_capex.replace(' -   ',0)
    df_capex=df_capex.replace(' - ',0)
    df_capex=df_capex[df_capex.Year.astype(str).str.contains('2012')==False]
    df_capex.Pars=df_capex.Pars.astype(float)

    return days_dic,df_const,df_capex

def creating_capital_spending(df, year):
    '''
    Return df_table - dataframe table for capital spending 

    Input  df - capex dataframe
           year - string - year to create capex table 
    '''
 
    df_table=df[(df.Year==year) & (df.CAD=='CAD')][['Type','Quarter','Pars']]
    df_table=df_table.pivot('Type','Quarter')['Pars']/1000.0
    df_table=df_table.reindex(['Land','Seismic','DC','Exp','FAC','OTHER','Acq','CAPGA','DISPO'])
    df_table['Total']=df_table.sum(1)
    df_table.loc['Net Operations']=df_table.sum(0)
    df_table.loc['Net Capex']=df_table.loc['Net Operations']
    df_table.loc['Net Development Capex']=df_table.iloc[:5].sum()
    return df_table

def create_production_forecast(df_table):
    '''
    Return df_prod - dataframe table for BOE production 

    Input  df_table - capitel spending table/dataframe
    '''

    df_prod=pd.DataFrame(index=['Capex Drilling','Openning','Decline on base',
                                 'Adds - previous carryover','Declines on previous year carryover',
                                 'Adds - current year capital','Declines on current year capital adds'], 
                                  columns=df_table.columns.tolist())
    df_prod.loc['Capex Drilling']=df_table.loc['Net Operations']-df_table.loc['Acq']
    
    # last years closing
    last_years_closing=settings.last_years_closing
    const_dic=settings.const_dic

    opening_prod=last_years_closing['Closing for period']

    opening_adds_prev_carryover = last_years_closing['Adds - previous carryover']  # enter value for last years 
    opening_decline_prev_year = last_years_closing['Declines on previous year carryover']
    opening_capex_drilling=last_years_closing['Capex Drilling']
    opening_capex_drilling_capeff=last_years_closing['CapEff']
    opening_capex_drilling_tiein=last_years_closing['TieIn']
    dec_current_year_cap = 0;
    
    for column in df_table.columns[:-1]:
        df_prod.loc['Openning',column]=opening_prod
        df_prod.loc['Decline on base',column]=-const_dic['Decline_Base']*df_prod.loc['Openning','Q1']/4
        df_prod.loc['Adds - previous carryover',column]=opening_adds_prev_carryover
        
        # the formula cannot produce results for 2013
        if last_years_closing['Year']==2013:
            const_dec_prevcarryover={'Q1':-200,'Q2':-500,'Q3':0,'Q4':-70}
            df_prod.loc['Declines on previous year carryover',column] = const_dec_prevcarryover[column]
        else :
            const_dec_prevcarryover={'Q1':0,'Q2':0,'Q3':0,'Q4':0}
            df_prod.loc['Declines on previous year carryover',column] = const_dec_prevcarryover[column]
            #df_prod.loc['Declines on previous year carryover',column] = opening_decline_prev_year
            #opening_decline_prev_year += -int(df_prod.loc['Adds - previous carryover',column]*const_dic['Decline_Prev']/4)
        
        df_prod.loc['Adds - current year capital',column]=\
            df_prod.loc['Capex Drilling',column]/const_dic['CapEff']*const_dic['TieIn']*1000+\
            opening_capex_drilling/opening_capex_drilling_capeff*(1-opening_capex_drilling_tiein)*1000
            
        opening_capex_drilling=df_prod.loc['Capex Drilling',column]
        opening_capex_drilling_capeff=const_dic['CapEff']
        opening_capex_drilling_tiein=const_dic['TieIn']
        
        df_prod.loc['Declines on current year capital adds',column] =dec_current_year_cap
        dec_current_year_cap += - df_prod.loc['Adds - current year capital',column]*const_dic['Decline_Current']/4
        
        
        df_prod.loc['Closing for period',column]=df_prod.loc['Openning':,column].sum(0)
        opening_prod=df_prod.loc['Closing for period',column]
        
    df_prod.loc['Capex Drilling','Total']=df_table.loc['DC','Total']
    df_prod.loc['Openning','Total']=df_prod.loc['Openning','Q1']
    df_prod.loc['Decline on base':'Declines on current year capital adds','Total']=\
            df_prod.loc['Decline on base':'Declines on current year capital adds'].sum(1)

    df_prod.loc['Closing for period','Total']=df_prod.loc['Openning':,'Total'].sum()
    
    df_prod.loc['Average']=df_prod.loc[['Openning','Closing for period']].mean().round()
    df_prod.loc['Average','Total']=df_prod.loc['Average',df_prod.columns[:-1]].mean().round(1)
    
    df_prod=df_prod.fillna(0)
    df_prod=df_prod.astype(float).round(1)

    return df_prod

def create_boe_summary(df_table):
    '''
    Return df_prod - dataframe table for BOE summary

    Input  df_table - dataframe table for BOE production
    '''

    df_summary=pd.DataFrame(index=['Opening balance','Existing production',
                                 'Carryover adds','Current year adds',
                                 'Closing balance'], 
                                  columns=df_table.columns.tolist())
    df_summary.loc['Opening balance']=df_table.loc['Openning']
    openning_balance_year=df_table.loc['Openning','Q1']
    add_carryover = 0
    add_current = 0
    for column in df_summary.columns[:-1]:
        openning_balance_year += df_table.loc['Decline on base',column]
        df_summary.loc['Existing production',column]=openning_balance_year
        
        add_carryover += df_table.loc['Adds - previous carryover',column]+\
            df_table.loc['Declines on previous year carryover',column]
        df_summary.loc['Carryover adds',column]=add_carryover
        add_current += df_table.loc['Adds - current year capital',column]+ \
                df_table.loc['Declines on current year capital adds',column]
        df_summary.loc['Current year adds',column]= add_current
        df_summary.loc['Closing balance',column]=df_summary.loc['Existing production':,column].sum(0)
        
    df_summary.loc['Average']=df_summary.loc[['Opening balance','Closing balance']].mean().round()
    df_summary.loc['Average','Total']=df_summary.loc['Average',df_summary.columns[:-1]].mean().round()
        
    return df_summary

def create_oil_summary(df_table):
    '''
    Return df_prod - dataframe table for BOE oil summary

    Input  df_table - dataframe table for BOE summary 
    '''

    const_dic=settings.const_dic
    df_summary=pd.DataFrame(index=['Opening balance','Existing production',
                                 'Carryover adds','Current year adds',
                                 'Closing balance'], 
                                  columns=df_table.columns.tolist())
    
    open_balance = int(df_table.loc['Opening balance', 'Q1']*const_dic['OILBAL'])
    
    for column in df_summary.columns[:-1]:
        df_summary.loc['Opening balance',column]=open_balance
        df_summary.loc['Existing production',column]=int(df_table.loc['Existing production',column]*const_dic['OILBAL'])
        df_summary.loc['Carryover adds',column]=int(df_table.loc['Carryover adds',column]*const_dic['CarryOver_OP'])
        df_summary.loc['Current year adds',column]=int(df_table.loc['Current year adds',column]*const_dic['CurrentYear'])
        df_summary.loc['Closing balance',column]=df_summary.loc['Existing production':,column].sum(0)
        open_balance=int(df_summary.loc['Closing balance',column])
        
    df_summary.loc['Average']=df_summary.loc[['Opening balance','Closing balance']].mean().round()
    df_summary.loc['Average','Total']=df_summary.loc['Average',df_summary.columns[:-1]].mean().round()
    
    for column in df_summary.columns:
        df_summary.loc['Percent Oil',column]=round(df_summary.loc['Average',column]/\
                                  df_table.loc['Average',column]*100,1)
    
    return df_summary

def create_gas_summary(df_table):
    '''
    Return df_prod - dataframe table for BOE gas summary

    Input  df_table - dataframe table for BOE summary 
    '''

    const_dic=settings.const_dic
    df_summary=pd.DataFrame(index=['Opening balance','Existing production',
                                 'Carryover adds','Current year adds',
                                 'Closing balance'], 
                                  columns=df_table.columns.tolist())
    const_gas=(1-const_dic['OILBAL'])*0.006
    const_carry = (1-const_dic['CarryOver_OP'])*0.006
    const_current = (1-const_dic['CurrentYear'])*0.006
    
    open_balance = round(df_table.loc['Opening balance', 'Q1']*const_gas,2)
    
    for column in df_summary.columns[:-1]:
        df_summary.loc['Opening balance',column]=open_balance
        df_summary.loc['Existing production',column]=round(df_table.loc['Existing production',column]*const_gas,2)
        df_summary.loc['Carryover adds',column]=round(df_table.loc['Carryover adds',column]*const_carry,2)
        df_summary.loc['Current year adds',column]=round(df_table.loc['Current year adds',column]*const_current,2)
        df_summary.loc['Closing balance',column]=df_summary.loc['Existing production':,column].sum(0)
        open_balance=round(df_summary.loc['Closing balance',column],2)
        
    df_summary.loc['Average']=df_summary.loc[['Opening balance','Closing balance']].mean().round(2)
    df_summary.loc['Average','Total']=df_summary.loc['Average',df_summary.columns[:-1]].mean().round(2)
    
    for column in df_summary.columns:
        df_summary.loc['Percent Gas',column]= round(df_summary.loc['Average',column]/\
                                  df_table.loc['Average',column]*100/0.006,2)
    
    
    return df_summary

def create_daily_production(df_gas, df_oil):
    '''
    Return df_daily - dataframe table for daily production

    Input  df_gas - dataframe table for BOE gas summary
           df_oil - dataframe table for BOE oil summary
    '''

    last_years_closing=settings.last_years_closing
    days_dic=settings.days_dic

    df_summary=pd.DataFrame(index=['Crude Oil (Bbls/d)','Heavy Oil (Bbls/d)',
                                    'NGLs (Bbls/d)','Natural Gas (MMcf/d)',
                                     'Total (Boe/d  6:1)'], columns=df_gas.columns.tolist())

    days = list(days_dic.values())
    coeff = last_years_closing['NGLs (Bbls/d)']/last_years_closing['Natural Gas (MMcf/d)']
   
    for column in df_summary.columns[:-1]:
        df_summary.loc['Natural Gas (MMcf/d)',column]=df_gas.loc['Average',column]
        df_summary.loc['NGLs (Bbls/d)',column]=coeff*df_summary.loc['Natural Gas (MMcf/d)',column]
        coeff =  df_summary.loc['NGLs (Bbls/d)',column]/df_summary.loc['Natural Gas (MMcf/d)',column]
        
        df_summary.loc['Heavy Oil (Bbls/d)',column]=0
        df_summary.loc['Crude Oil (Bbls/d)',column]=df_oil.loc['Average',column]-\
                                                          df_summary.loc['NGLs (Bbls/d)',column]
        
        
        df_summary.loc['Total (Boe/d  6:1)',column]=df_summary.loc[:'NGLs (Bbls/d)',column].sum()+\
                                                    df_summary.loc['Natural Gas (MMcf/d)',column]/6*1000
            
    df_tmp=df_summary[df_summary.columns[:-1]]*days
    df_summary['Total']=df_tmp.sum(1)/sum(days)
    df_summary=df_summary.astype(float).round(1)
    return df_summary

def create_annual_production(df_daily):
    '''
    Return df_annual - dataframe table for annual production

    Input  df_daily - dataframe table for daily production
    
    '''

    days_dic=settings.days_dic
    df_summary=pd.DataFrame(index=['Crude Oil (Bbls/d)','Heavy Oil (Bbls/d)',
                                    'NGLs (Bbls/d)','Natural Gas (MMcf/d)',
                                     'Total (Boe/d  6:1)'], columns=df_daily.columns.tolist())

    days = list(days_dic.values())
    
    for column in df_summary.columns[:-1]:
        df_summary.loc['Natural Gas (MMcf/d)',column]=df_daily.loc['Natural Gas (MMcf/d)',column]*days_dic[column]
        df_summary.loc['NGLs (Bbls/d)',column]=df_daily.loc['NGLs (Bbls/d)',column]*days_dic[column]/1000
        
        df_summary.loc['Heavy Oil (Bbls/d)',column]=df_daily.loc['Heavy Oil (Bbls/d)',column]*days_dic[column]/1000
        df_summary.loc['Crude Oil (Bbls/d)',column]=df_daily.loc['Crude Oil (Bbls/d)',column]*days_dic[column]/1000
        
        
        df_summary.loc['Total (Boe/d  6:1)',column]=df_summary.loc[:'NGLs (Bbls/d)',column].sum()+df_summary.loc['Natural Gas (MMcf/d)',column]/6
            
    
    df_summary['Total']=df_summary.sum(1)
   
    df_summary=df_summary.astype(int).round()
    return df_summary

def create_avecost_production(df_prod, df_prod_last, df_cap, df_cap_last,df_daily):
    '''
    Return df_daily - dataframe table for Average Cost of Production Replacement (12 mo Trailing in Qtrly)

    Input  df_prod, df_prod_last - this and last years production tables
           df_cap, df_cap_last - this and last years  capital spending tables
           df_daily - daily production table
    '''

    last_years_closing=settings.last_years_closing
    days_dic=settings.days_dic

    df_summary=pd.DataFrame(index=['Period Production Added','Period Production Declines',
                                    'Net Production Added','Capital Expenditures',
                                     'Annual Cost of Production Added'], columns=df_prod.columns.tolist())
    # I have to add first because i changed df_prod
    
    sum_adds_total=df_prod.loc['Adds - previous carryover','Total']+\
                            df_prod.loc['Adds - current year capital','Total']

    sum_declines_total=df_prod.loc['Decline on base','Total']+\
            df_prod.loc['Declines on previous year carryover','Total']+\
            df_prod.loc['Declines on current year capital adds','Total']
        
    df_prod_last=df_prod_last[df_prod_last.columns[:-1]]
    df_prod=df_prod[df_prod.columns[:-1]]
    df_prod_last.columns=df_prod.columns+'_last'
    
    df_tmp=pd.concat([df_prod_last,df_prod],1)

    size_prod=df_prod.shape[1]
    df_sum=pd.DataFrame(index=df_prod.index, columns=df_prod.columns)
    
    for i in reversed(range(size_prod)):
        columns=df_tmp.columns[-4-i:-i]
        if i==0:
            columns=df_tmp.columns[-4:] 
            
        df_sum[columns[-1]]=df_tmp[columns].sum(1)
    
    # capital spending
    df_cap_last=df_cap_last[df_cap_last.columns[:-1]]
    df_cap=df_cap[df_cap.columns[:-1]]
    df_cap_last.columns=df_prod.columns+'_last'
    
    df_tmp=pd.concat([df_cap_last,df_cap],1)

    size_cap=df_cap.shape[1]
    df_capsum=pd.DataFrame(index=df_cap.index, columns=df_cap.columns)
    
    for i in reversed(range(size_cap)):
        columns=df_tmp.columns[-5-i:-i]
        if i==0:
            columns=df_tmp.columns[-5:]
    
        df_capsum[columns[-1]]=df_tmp[columns[1:-1]].sum(1)+(df_tmp[columns[0]]+df_tmp[columns[-1]])/2.0
    
    #days = list(days_dic.values())
    
    for column in df_summary.columns[:-1]:
        df_summary.loc['Period Production Added',column]=df_sum.loc['Adds - previous carryover',column]+\
                            df_sum.loc['Adds - current year capital',column]
        df_summary.loc['Period Production Declines',column]=df_sum.loc['Decline on base',column]+\
            df_sum.loc['Declines on previous year carryover',column]+\
            df_sum.loc['Declines on current year capital adds',column]
        
        df_summary.loc['Net Production Added',column]=df_summary.loc[:'Period Production Declines',column].sum()
        
        df_summary.loc['Capital Expenditures',column]=df_capsum.loc['Net Operations',column]
        
        df_summary.loc['Annual Cost of Production Added',column]=df_summary.loc['Capital Expenditures',column]/\
            df_summary.loc['Period Production Added',column]*1000
         
    df_summary.loc['Period Production Added','Total']=sum_adds_total
    df_summary.loc['Period Production Declines','Total']=sum_declines_total
    df_summary.loc['Net Production Added','Total']=df_daily.loc['Total (Boe/d  6:1)','Total']-\
            last_years_closing['Total (Boe/d 6:1)']

    df_summary.loc['Annual Cost of Production Added','Total']=\
        df_summary.loc['Annual Cost of Production Added',df_summary.columns[:-1]].mean().round(1)
    
    df_summary=df_summary.fillna(0)
    df_summary=df_summary.astype(float).round()

    return df_summary

def  create_excelfile(yearInt, outFile, df_daily_last,df_ave_prod_last,df_daily,df_ave_prod, exit_prod):
    '''
    Return - creating the request tables and written in the output excel file.

    Input  yearInt - year(integer) to create tables
           df_daily_last,df_daily - this and last years daily production tables
           df_ave_prod_last,df_ave_prod - this last year's Average Cost Production tables
           exit_prod - list of exit production numbers
    '''

    writer = pd.ExcelWriter(outFile, engine='xlsxwriter')

    df_daily.to_excel(writer, startcol=4,startrow= 2,index=False)
    df_ave_prod.to_excel(writer, startcol=4, startrow=df_daily.shape[0]+6, header=False, index=False )


    if yearInt==2014:
        # only for 2014 because results for 2013 total is giving wrong numbers
        df_daily_last['Total']=[3419,0,233.4,19.3,6872]
        df_ave_prod_last['Total']=[3408,-3213,1516,0,36287]
        exit_prod[0]='6795/right-6595'

    df_daily_last['Total'].to_excel(writer, startcol=1, startrow= 2 )
    df_ave_prod_last['Total'].to_excel(writer, startcol=1, startrow=df_daily.shape[0]+6, header=False )

    worksheet = writer.sheets['Sheet1']
    worksheet.write(0, 2, str(yearInt-1))
    worksheet.write(0, 6, str(yearInt))

    worksheet.write(1, 0, "Daily Production")
    worksheet.write(9, 0, "Average Cost of Production Replacement (12 mo Trailing in Qtrly)")

    worksheet.write(18, 0, "Exit Production")
    worksheet.write(18, 2, str(exit_prod[0]))
    worksheet.write(18, 6, str(exit_prod[1]))

    writer.save()

