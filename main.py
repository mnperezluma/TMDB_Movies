#!/usr/bin/env python
# coding: utf-8



import requests
import pandas as pd
import numpy as np
import warnings
import smtplib
import win32com.client as win32
import os
from datetime import date
outlook = win32.Dispatch("Outlook.Application")
warnings.filterwarnings("ignore")


# Credentials
API_KEY = os.getenv("SECRET_API_KEY")




def get_popular_movies():
    # Function which brings the most popular movies from every country

    query_pop =  'https://api.themoviedb.org/3/movie/popular?api_key={}&language=en-US&page=1'.format(API_KEY)
    ans = requests.get(query_pop)
    POPULAR = ans.json()
    
    
    try:
        LISTA_POPULAR = []
        
        for movie in POPULAR['results']:
            pop_tuple = (movie['id'],movie['title'], movie['original_language'],movie['vote_average'], movie['popularity'])
            LISTA_POPULAR.append(pop_tuple)
            
        df = pd.DataFrame.from_records(LISTA_POPULAR, columns=['ID', 'Title', 'Original_language', 'Average_vote', 'Popularity'])
        
        return df
    except Exception as error:
        print(error)
        return -1

def check_availability_country(country:str):
    # Checking the availability of that country
    # Parameters:
    # country: Name of the country (str)

    COUNTRIES ="https://api.themoviedb.org/3/configuration/countries?api_key={}".format(API_KEY)
    req = requests.get(COUNTRIES)
    resp = req.json()
    LIST_COUNTRIES = [list([resp[x]['english_name'],resp[x]['iso_3166_1']]) for x in range(len(resp))]
    DC_COUNTRIES = dict(LIST_COUNTRIES)
    
    try:
        if country in DC_COUNTRIES.keys():
            AB = DC_COUNTRIES[country]
            return AB
        else:
            print("The country you chose is not available. Please choose another one or recheck your typing")
            return None
    except Exception as e:
        print("Please re-try later")




def get_provider(country:str):
    # Function which informs TV provider for each one of the top movies
    # Parameters:
    # country: Name of the country (str)

    df = get_popular_movies()

    if isinstance(df, pd.DataFrame):
        df['TV Provider'] = 0
        country = check_availability_country(country)
        try:
            for movie_id in df['ID']:
                PROVIDERS = 'https://api.themoviedb.org/3/movie/{}/watch/providers?api_key={}'.format(movie_id,API_KEY)
                req = requests.get(PROVIDERS)
                resp = req.json()
                if country in resp['results'].keys():
                    try:
                        df['TV Provider'] = np.where(df['ID'] == movie_id, resp['results'][country]['flatrate'][0]['provider_name'],df['TV Provider'])
                    except:
                        df['TV Provider'] = np.where(df['ID'] == movie_id, 'No Flatrate available',df['TV Provider'])

                else:
                    df['TV Provider'] = np.where(df['ID'] == movie_id, 'Not available in {}'.format(country),df['TV Provider'])
            
            df.to_csv('Reports/Report_{}.csv'.format(date.today()), index=False)
            return df
        except Exception as Error:
            print(Error)
    else:
        print("There is an error extracting data from API")
        return None



def send_report(country:str, email:str, df:pd.DataFrame):
    # Function which sents a report to e-mail provided
    # Parameters:
    # country: Name of the country (str)
    # email: E-mail provided by the user(str)
    # df: Report that will be sent including top 20 popular movies (pd.DataFrame)


    try:

        country_select = country
        mail = outlook.CreateItem(0)
        mail.To = email
        mail.Subject = 'Popular Movies Report {}'.format(country_select)
        

        mail.HTMLBody = '''    <html>
        <head>
        <style>
        body {background-color: powderblue;}
        h1   {color: blue;}
        p    {color: red;}
        </style>
        </head>
        <body>

        <h1>Top popular Movies</h1>

        ''' + country_select + df.to_html() + '''


        </body>
        </html>'''

        mail.Send()
        print("The report has been sent")
    except Exception as error:
        print(error)
    


if __name__ == '__main__':    
    country_selection = input("Please choose the country where you want to receive the report: ")
    mail_selection = input("Please include your e-mail adress where you want to receive the report: ")
    if not isinstance(country_selection, str) and (mail_selection,str):
        print("Both inputs must be a string. Re check it. ")
    
    df_final = get_provider(country_selection)
    send_report(country_selection,mail_selection, df_final)
    





