import requests
import lxml.html as lh
import pandas as pd
def webscrape_df(url, num_col):
    #reading data from url
    page = requests.get(url)
    doc = lh.fromstring(page.content)
    #Parse data that are stored between <tr>..</tr> of HTML
    tr_elements = doc.xpath('//tr')

    col=[]
    i=0
    #header
    for t in tr_elements[0]:
        i+=1
        name=t.text_content()
        col.append((name,[]))

    #data
    for j in range(1,len(tr_elements)):
        #T is the j'th row of the table
        T=tr_elements[j]
        
        #If row is not of size 12, the //tr data is not from the table we want
        #just to check
        if len(T)!=num_col:
            break
        
        #column index
        i=0
        
        #Iterate through each element of the row
        for t in T.iterchildren():
            data=t.text_content() 
            #Append the data to the empty list of the i'th column
            col[i][1].append(data)
            #Increment i for the next column
            i+=1

    #create dictionary with column name as key and a list of the column's contents as value
    Dict={title:column for (title,column) in col}
    df=pd.DataFrame(Dict)
    return df
