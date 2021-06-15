from selenium import webdriver
import time
import numpy as np
from webdriver_manager.chrome import ChromeDriverManager
import base64
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (Mail, Attachment, FileContent, FileName, FileType, Disposition)
from string import Template
from datetime import date, timedelta
import os
import pandas as pd
from metabasepy import Client, MetabaseTableParser

SENDGRID_API_KEY = os.getenv('SENDGRID_API_KEY')
user = os.getenv('USER')
password = os.environ.get('PASSWORD')
endpoint = "http://metabase.yourcompany.tech"
yesterday = (date.today() - timedelta(1)).strftime('%d-%m-%Y')
yesterday_f = (date.today() - timedelta(1)).strftime('%Y-%m-%d')
url = "http:your.dashboard.url?date=%s" % yesterday_f
mess = open('dailyreport.html', 'r', encoding='utf-8').read()


def run():
    print("Here we go...")
    start = time.time()
    out = get_metabase_data(user, password, endpoint, url, yesterday, SENDGRID_API_KEY)
    take_screenshot(user, password, url)
    send_mail(out[0], out[1], out[2], out[3], out[4], out[5], out[6], out[7], out[8], out[9], out[10], out[11], out[12],
              out[13], out[14], out[15], out[16], out[17], out[18], out[19], out[20], out[21], mess)
    print("I'm done...I'm going back to bed...")
    end = time.time()
    print('execution time of script: ', str(timedelta(seconds=(end - start))))


def get_metabase_data(user, password, endpoint, url, yesterday, SENDGRID_API_KEY):
    cli = Client(username=user, password=password, base_url=endpoint)
    cli.authenticate()
    query_response = cli.cards.query(card_id="708")
    print(query_response)
    data_table = MetabaseTableParser.get_table(metabase_response=query_response)
    df = pd.DataFrame(data_table.__dict__['rows'])
    df.columns = ['date', 'orders', 'qty', 'gp', 'inc', 'aio', 'aov']
    print(df)

    orders = df['orders'][0]
    items = df['qty'][0]
    gp = df['gp'][0]
    income = df['inc'][0]
    aio = df['aio'][0]
    aov = df['aov'][0]

    query_response = cli.cards.query(card_id="720")
    print(query_response)
    data_table = MetabaseTableParser.get_table(metabase_response=query_response)
    df1 = pd.DataFrame(data_table.__dict__['rows'])
    df1.columns = ['date', 'orders', 'qty', 'gp', 'inc', 'aio', 'aov']
    print(df1)

    ordersil = df1['orders'][0]
    itemsil = df1['qty'][0]
    gpil = df1['gp'][0]
    incomeil = df1['inc'][0]
    aioil = df1['aio'][0]
    aovil = df1['aov'][0]

    query_response = cli.cards.query(card_id="721")
    print(query_response)
    data_table = MetabaseTableParser.get_table(metabase_response=query_response)
    df2 = pd.DataFrame(data_table.__dict__['rows'])
    df2.columns = ['date', 'orders', 'qty', 'gp', 'inc', 'aio', 'aov']
    print(df2)

    ordersus = df2['orders'][0]
    itemsus = df2['qty'][0]
    gpus = df2['gp'][0]
    incomeus = df2['inc'][0]
    aious = df2['aio'][0]
    aovus = df2['aov'][0]

    query_response = cli.cards.query(card_id="709")
    print(query_response)
    data_table = MetabaseTableParser.get_table(metabase_response=query_response)
    df3 = pd.DataFrame(data_table.__dict__['rows'])
    df3.columns = ['category', 'itemname', 'itemcolor', 'asn', 'sales_total_qty', 'income_total', 'sales_usweb_qty',
                   'sales_ilweb_qty', 'sales_ilstores_qty']
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    df3.index = np.arange(1, len(df3) + 1)
    df3 = df3.to_html(index_names=False)

    lst = [ordersil, itemsil, gpil, incomeil, aioil, aovil,
           ordersus, itemsus, gpus, incomeus, aious, aovus, orders, items, gp, income, aio, aov, df3, url, yesterday,
           SENDGRID_API_KEY]
    return lst


def take_screenshot(user, password, url):
    options = webdriver.ChromeOptions()
    options.add_argument('--no-sandbox')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument("--disable-dev-shm-usage");
    # Install Driver and change window size
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    driver.set_window_size(1920, 1080, driver.window_handles[0])

    # # Specify URL
    print(url)
    # # # Open URL
    driver.get(url)
    # #
    # # # Click on "Sign in with email"
    driver.find_element_by_link_text("Sign in with email").click()
    #
    # # Identify username, password and sign-in elements
    print("Entering Metabase")
    driver.find_element_by_name("username").send_keys(user)
    time.sleep(0.2)
    driver.find_element_by_name("password").send_keys(password)
    time.sleep(0.4)
    driver.find_element_by_css_selector(
        "#root > div > div > div > div > form > div.Form-actions.text-centered > button > div").click()
    time.sleep(10)

    # # # Open the dashboard
    print("Opening folder 1")
    driver.find_element_by_css_selector(
        "#root > div > div.sc-bdVaJa.iHZvIS > div.sc-bdVaJa.bhVfai > div.sc-bdVaJa.dpyHrh > div > div > div:nth-child(1) > div > div:nth-child(1) > a > div > div").click()
    time.sleep(5)
    print("Opening folder 2")
    driver.find_element_by_css_selector(
        "#root > div > div.Layout__PageWrapper-sc-3nh584-0.imuuiX.sc-bdVaJa.iHZvIS > div.CollectionSidebar__Sidebar-ktzkk8-0.gvNhBz.sc-bdVaJa.fMagvA > div.sc-bdVaJa.kEyBEd > div:nth-child(1) > div:nth-child(1) > div.sc-bdVaJa.dpqWwQ > div > div:nth-child(1) > div > a > div").click()
    time.sleep(5)
    print("Opening folder 3")
    driver.find_element_by_css_selector(
        "#root > div > div.Layout__PageWrapper-sc-3nh584-0.imuuiX.sc-bdVaJa.iHZvIS > div.border-left.full-height.sc-bdVaJa.klotAC > div.sc-bdVaJa.eMAcAT > div.sc-bdVaJa.gwldrd > div.relative.sc-bdVaJa.iHZvIS > div.relative > div.sc-bdVaJa.iHZvIS > div:nth-child(1) > div > div > div:nth-child(1) > div > div > a > div > div.sc-bdVaJa.iHZvIS > h3 > div").click()
    time.sleep(30)

    # # Take a screenshot
    print("Taking a screenshot")
    S = lambda x: driver.execute_script('return document.body.parentNode.scroll' + x)
    driver.set_window_size(S('Width'), S('Height'))
    driver.find_element_by_css_selector(
        "#root > div > div.shrink-below-content-size.full-height > div > div > div").screenshot('daily report.png')
    driver.close()  # close chrome window
    print("Screenshot taken successfully")


def send_mail(ordersil, itemsil, gpil, incomeil, aioil ,aovil, ordersus, itemsus, gpus, incomeus, aious, aovus,
              orders, items, gp, income, aio, aov, df3, url, yesterday, SENDGRID_API_KEY, mess):
    with open('daily report.png', 'rb') as f:
        data = f.read()
        f.close()
    encoded_file = base64.b64encode(data).decode()
    mess = Template(mess).substitute(url=url, yesterday=yesterday, orders=orders,
                                     items=items, gp=gp, income=income, aio=aio, aov=aov,
                                     ordersil=ordersil, itemsil=itemsil, gpil=gpil, incomeil=incomeil, aioil=aioil,
                                     aovil=aovil, ordersus=ordersus, itemsus=itemsus, gpus=gpus,
                                     incomeus=incomeus, aious=aious, aovus=aovus, df=df3)

    print("Trying to send email")
    message = Mail(
        from_email='your@email.com',
        to_emails='forward@email.com',
        subject="Daily Report %s " % yesterday,
        plain_text_content="encoded_file_a",
        html_content=mess
    )

    attachedFile = Attachment(
        FileContent(encoded_file),
        FileName('attachment.png'),
        FileType('application/png'),
        Disposition('attachment'),
    )
    message.attachment = attachedFile

    sg = SendGridAPIClient(api_key=SENDGRID_API_KEY)
    response = sg.send(message)
    print(response.status_code, response.body, response.headers)
    print("Mail sent successfully")



