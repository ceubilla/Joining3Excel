import pandas as pd
import pandas_gbq as gbq

def read_xlsxdata(filename):
    dataset = pd.read_excel(filename,sheet_name='Sheet1') 
    return dataset

# Read data from excel files
dataset_trebas = read_xlsxdata('TREBAS.xlsx') 
dataset_mcgill = read_xlsxdata('McGill.xlsx')
dataset_concordia = read_xlsxdata('Concordia.xlsx')

#fixing the column names in trebas dataset
dataset_trebas['Full_Name'] = dataset_trebas['First_Name'] + ' ' + dataset_trebas['Last_Name']

#Adding a column to the trebas dataset to identify the university
dataset_trebas['University'] = 'TREBAS INSTITUTE MONTREAL'

#Adding a column to the mcgill dataset to identify the university
dataset_mcgill['University'] = 'McGill University'

#Adding a column to the concordia dataset to identify the university
dataset_concordia['University'] = 'Concordia University'

#Creating dataframes for each university
df_trebas = pd.DataFrame(dataset_trebas, columns = ['Full_Name','Country','Program','University'])
df_mcgill = pd.DataFrame(dataset_mcgill, columns = ['name','country','Course','University'])
df_concordia = pd.DataFrame(dataset_concordia, columns = ['Full_Name','Nationality','Program','University'])

#Setting the column names to be the same for all dataframes
df_mcgill.rename(columns = {'name':'Full_Name'},inplace = True)
df_mcgill.rename(columns = {'country':'Country'},inplace = True)
df_mcgill.rename(columns = {'Course':'Program'},inplace = True)
df_concordia.rename(columns = {'Nationality':'Country'},inplace = True)

#Concatenating all dataframes into one
df_universities = pd.concat([df_trebas,df_mcgill,df_concordia],ignore_index=True)

#Writing the dataframe to an Excel file

writer = pd.ExcelWriter('Universities.xlsx', engine='xlsxwriter')
df_universities.to_excel(writer, sheet_name='all_union')
writer.save()

#Writing the dataframe to BigQuery
""""
mylocation = 'northamerica-northeast2'
myproject = 'braided-topic-362415'
gbq.to_gbq(df_universities, 'universities_DB.intstudents',location=mylocation, project_id=myproject, if_exists='replace')
"""
