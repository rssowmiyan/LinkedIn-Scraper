from selenium.webdriver import Firefox
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
import time
from bs4 import BeautifulSoup
import xlsxwriter
import os 
import re
from selenium.webdriver.common.by import By
from pprint import pprint
# ************************************************************
# Use sessions to open firefox latest update
profile_path = r'C:\Users\Administrator\AppData\Roaming\Mozilla\Firefox\Profiles\y1uqp5mi.default'
options=Options()
options.set_preference('profile', profile_path)
service = Service(r'C:\GeckoDriver\geckodriver.exe')
driver = Firefox(service=service, options=options)
password = os.environ.get('password')
driver.get('https://www.linkedin.com/login')
driver.find_element_by_id('username').send_keys('sowmiyan00@gmail.com') 
driver.find_element_by_id('password').send_keys(password)
driver.find_element(by=By.XPATH,value="//*[@type='submit']").click()

# Go to college people section
search_url='https://www.linkedin.com/school/thiagarajar-college-of-engineering/people/'
driver.get(search_url)
driver.maximize_window()
data=[]

# Giving time for machine to scroll through the content
SCROLL_PAUSE_TIME = 20
last_height = driver.execute_script("return document.body.scrollHeight")

# Counter variable to keep track of full page scrolls
cnt=0
while True:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(SCROLL_PAUSE_TIME)
    new_height = driver.execute_script("return document.body.scrollHeight")
    if new_height == last_height:
        break
    last_height = new_height
    cnt+=1
    if(cnt==2):
        break

# lxml -> parser used here
search = BeautifulSoup(driver.page_source,'lxml')
peoples_div_section = search.findAll('div', attrs ={'class':'org-people-profile-card__profile-info'})
peoples_div_section = str(peoples_div_section)
# regex to filter the people <a href="value we need">
peoples = re.findall(r'href=[\'"]?([^\'" >]+)', peoples_div_section)
pprint(peoples)
count = 0
                
for people in peoples:
    count+=1 
    profile_url = "https://www.linkedin.com" + people
    print(profile_url)
    loc = ""
    page = BeautifulSoup(driver.page_source,'lxml')

    # title
    try:
        title = str(page.find("h1", attrs = {'class':'text-heading-xlarge inline t-24 v-align-middle break-words'}).text).strip()
        
    except:
        title = 'None'


    #heading
    try:
        heading = ''
        for temp in (page.select("div[class='text-body-medium break-words']")):
            temp=str(temp)
            pattern = re.compile(r'[^<>]+(?=.<)')
            heading = pattern.findall(temp)
            heading=''.join(heading)
        
        if(len(heading)==0):
            heading = None
        print("heading",heading)
                
    except:
        heading = 'None'


    # location
    try:
        loc = str(page.find('span', attrs = {'class':'text-body-small inline t-black--light break-words'}).text).strip()
        loc=loc.split(',')
        if(len(loc)==1):
            country=loc[0]
            state='None'
            district='None'
        elif(len(loc)==2):
            country=loc[1]
            state='None'
            district=loc[0]
        else:
            district=loc[0]
            state=loc[1]
            country=loc[2]
    except:
            country='None'
            state='None'
            district='None'
            
    try:
        connections=str(page.find('span', attrs = {'class':'t-bold'}).text).strip()
    except:
        connections='None'
        
    try:
        exp_section = exp_section.find('ul',attrs={'class':'pv-profile-section__section-info'})
        li_tags = exp_section.find('div')
        a_tags = li_tags.find('a')
    except:
        exps_section=page.select("ul[class='pv-profile-section__section-info section-info pv-profile-section__section-info--has-no-more']")
        # print('exp_Section= ',exps_section)
        for x in exps_section:
            print('exp_section')
            exp_section=x
            break
        lis_tags=page.select("li[class='pv-entity__position-group-pager pv-profile-section__list-item ember-view']")
        for x in lis_tags:
            print('li_tags')
            li_tags=x
            break
        as_tags=page.select("a[class='full-width ember-view']")
        for x in as_tags:
            print('a_tags')
            a_tags=x
            break
        
                                    
    try:
        job_title=str(a_tags.find('h3').get_text()).strip()
        if('Company' in job_title):
            job_title=''
            job=[]
            for x in page.select("h3[class='t-14 t-black t-bold']"):
                job.append(x.find_all('span')[1].get_text().strip())
            job_title=','.join(job)
            company_name='Thiagarajar College of Engineering'


    except:
        job=[]
        for x in page.select("h3[class='t-14 t-black t-bold']"):
            job.append(x.find_all('span')[1].get_text().strip())
        job_title=','.join(job)
            
    try:
        company_name = str(a_tags.find_all('p')[1].get_text()).strip()
    except:
        company_name='None'
            
        # try:
        #     joining_date = str(a_tags.find_all('h4')[0].find_all('span')[1].get_text()).strip()
            
        # except:
        #     joining_date='None'
        
    try:
            # print(a_tags.find_all('h4')[1].find_all('span')[1].get_text())
        exp =str(a_tags.find_all('h4')[1].find_all('span')[1].get_text()).strip()
    except (NameError):
        exp = None

    try:
        # print(a_tags.find_all('h4')[0].find_all('span')[1].get_text())
        exp= str(a_tags.find_all('h4')[0].find_all('span')[1].get_text()).strip()
    
    except:
        pass


    try:
        edu_section = page.find('section', {'id': 'education-section'}).find('ul')
    except:
        pass


    try:
        institution = page.find('h3',attrs={"class":"pv-entity__school-name t-16 t-black t-bold"}).get_text().strip()
        # print('insti=',institution)
    except:
            institution='None'


    try:
        flag=True;l2=[];l3=[]
        for edu in page.find_all('span',attrs={"class":"pv-entity__comma-item"}):#iknowbitmessy
            if(flag):
                l2.append(edu.get_text())
                flag=False
            else:
                l3.append(edu.get_text())
                Flag=True
        for zzz in zip(l2,l3):
                degree_name=' '.join(zzz)
                
    except:
        degree_name=None
        
    try:
        passed_out= edu_section.find('p', {'class': 'pv-entity__dates t-14 t-black--light t-normal'}).find_all('span')[1].get_text().strip()
    except:
        passed_out='None'
        

    # Contact Information
    time.sleep(3)
    driver.get(profile_url + 'detail/contact-info/')

    info = BeautifulSoup(driver.page_source, 'lxml')
    details = info.findAll('section',attrs = {'class':'pv-contact-info__contact-type'})

    try:
        websites = details[1].findAll('a')
        for website in websites:
            website = website['href']
    except:
        website = 'None'


    try:
        number = details[2].find('span').text
        if(number.isnumeric()):
            phone=number
        else:
            phone='None'
    except:
        phone = 'None'


    try:
            email = str(details[3].find('a').text).strip()
    except:
            email = 'None'
        
    data.append({'profile_url':profile_url,'title':title,'heading':heading,'country':country,'state':state,'district':district,'institution':institution,'degree_name':degree_name,'passed_out':passed_out,'website':website,'phone':phone,'email':email,'connections':connections,'job_title':job_title,'company_name':company_name,'exp':exp})

print("!!!!!! Data scrapped !!!!!!")
    
print(f'No of pages scrolled ={count}')

driver.close()


                        
workbook = xlsxwriter.Workbook(os.path.join(os.path.dirname(os.path.abspath(__file__)),"linkedindataFinal.xlsx"))
worksheet = workbook.add_worksheet('Peoples')
bold = workbook.add_format({'bold': True})
worksheet.write(0,0,'profile_url',bold)
worksheet.write(0,1,'Name',bold)
worksheet.write(0,2,'Job',bold)
worksheet.write(0,3,'country',bold)
worksheet.write(0,4,'state',bold)
worksheet.write(0,5,'district',bold)
worksheet.write(0,6,'Educated - Institution',bold)
worksheet.write(0,7,'degree_name',bold)
# worksheet.write(0,8,'stream',bold)
worksheet.write(0,8,'passed_out',bold)
worksheet.write(0,9,'website/social media',bold)
worksheet.write(0,10,'phone',bold)
worksheet.write(0,11,'email',bold)
worksheet.write(0,12,'connections',bold)
worksheet.write(0,13,'Experienced as',bold)
worksheet.write(0,14,'working - company_name/Institution',bold)
# worksheet.write(0,16,'joining_date',bold)
worksheet.write(0,15,'years of experience',bold)
for i in range(1,len(data)+1):
           
    try:
        worksheet.write(i,0,data[i]['profile_url'])

    except:
        pass
    try:
         worksheet.write(i,1,data[i]['title'])
    except:
        pass

    try:
        worksheet.write(i,2,data[i]['heading'])
    except:
        pass
    try:
        worksheet.write(i,3,data[i]['country'])
    except:
        pass
    try:
        worksheet.write(i,4,data[i]['state'])
    except:
        pass
    try:
        worksheet.write(i,5,data[i]['district'])
    except:
        pass
    try:
        worksheet.write(i,6,data[i]['institution'])
    except:
        pass
    try:
        worksheet.write(i,7,data[i]['degree_name'])
    except:
        pass
  
    try:
        worksheet.write(i,8,data[i]['passed_out'])
    except:
        pass
    try:
        worksheet.write(i,9,data[i]['website'])
    except:
        pass
    try:
        worksheet.write(i,10,data[i]['phone'])
    except:
        pass
    try:
        worksheet.write(i,11,data[i]['email'])
    except:
        pass
    try:
         worksheet.write(i,12,data[i]['connections'])
    except:
        pass
    try:
         worksheet.write(i,13,data[i]['job_title'])
    except:
        pass
    try:
         worksheet.write(i,14,data[i]['company_name'])
    except:
        pass
    try:
         worksheet.write(i,15,data[i]['exp'])
    except:
        pass


workbook.close()


