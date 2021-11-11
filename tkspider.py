import scrapy

import datetime
import csv
import sys
import math

class TkSpider(scrapy.Spider):
    name = "tkspider"
    url_result=[]

    def start_requests(self):
        url = getattr(self, 'url', None)
        if url is not None:
            url = url
        category=["TYPE OF RACE",'DATE','COURSE','DISTANCE/Y','GOING','CLASS','POSITION','HORSE NAME','AGE','WEIGHT','STONES','POUNDS','ALL POUNDS','JCK ALNC','3YO ALLOWANCE','WON','SCORE','MARGIN','COMMENTS']
        with open('./tmp/resultcsv.csv','a') as writeurl:
             writer=csv.writer(writeurl)                
             writer.writerow([url])
             writer.writerow(category)
             writeurl.close()
        print("scraping started correctly")
        yield scrapy.Request(url, self.parse)

    def parse(self, response):
        course_header=response.xpath('//div[@class="rp-raceTimeCourseName"]')

        course_h1=course_header.xpath('.//h1[@class="rp-raceTimeCourseName__header"]')
        course_name=course_h1.xpath('.//a[contains(@class,"rp-raceTimeCourseName__name")]/text()').extract_first()
        course_name=course_name.strip()
        course_date=course_h1.xpath('.//span[@class="rp-raceTimeCourseName__date"]/text()').extract_first()
        course_date=course_date.strip()
        month=""
        course_date=course_date.split()
        month=course_date[1]
        day=course_date[0]

        if month=="Jan":
            month="1"
        elif month=="Feb":
            month="2"
        elif month=="Mar":
            month="3"
        elif month=="Apr":
            month="4"
        elif month=="May":
            month="5"
        elif month=="Jun":
            month="6"
        elif month=="Jul":
            month="7"
        elif month=="Aug":
            month="8"
        elif month=="Sep":
            month="9"
        elif month=="Oct":
            month="10"
        elif month=="Nov":
            month="11"
        elif month=="Dec":
            month="12"

        
        xdate=datetime.datetime(int(course_date[2]),int(month),int(course_date[0]))
 
        course_date=xdate.strftime("%d/%m/%Y")
          
        course_info=course_header.xpath('.//div[@class="rp-raceTimeCourseName__info"]')

        course_type_expand=course_info.xpath('.//h2[@class="rp-raceTimeCourseName__title"]/text()').extract_first()
        type01=course_type_expand.strip()
        if "Handicap" in type01:
            if "Apprentic" in type01:
                type01="Apprentice Handicap"
            elif "Nursery" in type01:
                type01="Nursery Handicap"
            else:
                type01="Handicap"
        elif "Stakes" in type01:
            if "Novice" in type01:
                type01="Novice Stakes"
            elif "Condition" in type01:
                type01="Condition Stackes"
            elif "Maiden" in type01:
                type01="Maiden Stakes"
            elif "Claiming" in type01:
                type01="Claiming Stakes"
            elif "Selling" in type01:
                type01="Selling Stakes"
            elif "Classified" in type01:
                type01="Classified Stakes"
            else:
                type01="Stakes"
        elif "Listed" in type01:
            type01="Listed Race"
        else:
            type01="Nonetype:"  
        course_info_middle=course_info.xpath('.//span[@class="rp-raceTimeCourseName__info_container"]')
        course_class=course_info_middle.xpath('.//span[@class="rp-raceTimeCourseName_class"]/text()').extract_first()
        course_class=course_class.strip().strip("(").strip(")")
        ageofhorse_expand=course_info_middle.xpath('.//span[@class="rp-raceTimeCourseName_ratingBandAndAgesAllowed"]/text()').extract_first()
        ageofhorse_expandlist=ageofhorse_expand.split(",")
        ageofhorselistvalue=ageofhorse_expandlist[-1:][0]
        ageofhorse=ageofhorselistvalue.strip()
        ageofhorse=ageofhorse.strip("(").strip(")")
        type_of_race=type01+" "+ageofhorse
        course_type02=course_info.xpath('.//span[@class="rp-raceTimeCourseName_distanceDetail"]/text()').extract_first()
        if course_type02==None:
            if course_name=="Nottingham":
                course_type02="Outer"
            elif course_name=="Doncaster":
                course_type02="Rnd"
            elif course_name=="Haydock":
                course_type02="Outer"
            elif course_name=="Newbury Str":
                course_type02="Rnd"
            elif course_name=="Newmarket":
                course_type02="Row"
            else:
                course_type02=""
        else:
            course_type02=course_type02.strip()
        course_name=course_name+' '+course_type02 
        course_going=course_info_middle.xpath('.//span[@class="rp-raceTimeCourseName_condition"]/text()').extract_first()        
        course_going=course_going.strip()
  
        maindistance=course_info_middle.xpath('.//span[@class="rp-raceTimeCourseName_distanceFull"]/text()').extract_first()
 
        if maindistance:

            maindistance=maindistance.strip()
            maindistance=maindistance.strip("(").strip(")")
            maindistance=list(maindistance)
            mm=""
            distance_yard=0
            for maindisvalue in maindistance:
                  
                if maindisvalue=="m":

                    distance_yard=8*220*int(mm)
                    mm=""
                    continue
                if maindisvalue=="f":
   
                    distance_yard=distance_yard+220*int(mm)
                    mm=""
                    continue
                if maindisvalue=="y":
  
                    distance_yard=distance_yard+int(mm)
                mm=mm+maindisvalue
  
        else:            
            course_distance=course_info_middle.xpath('.//span[@class="rp-raceTimeCourseName_distance"]/text()').extract_first()
            course_distance=course_distance.strip()
  
            general_distance=list(course_distance)
 
            mm=""
            distance_yard=0
            for generalvalue in general_distance:
                   
                if generalvalue=="m":
       
                    distance_yard=8*220*int(mm)
                    mm=""
                    continue
                if generalvalue=="f":
  
                    ff=list(mm)
                    if ff[-1:][0]=="½":                    
                        distance_yard=distance_yard+220*int(ff[:-1][0])+110
                    else:
                        distance_yard=distance_yard+220*int(mm)
                    mm=""
                    continue
                if generalvalue=="y":
                    distance_yard=distance_yard+int(mm)
                mm=mm+generalvalue
        
        mainrow_list=response.xpath('.//tr[@class="rp-horseTable__mainRow"]')      
        resultlist=[]
        marginvalue_list=[]
        for mainrow in mainrow_list:
            won=""
            gamerank=mainrow.xpath('.//td[1]/div/div/span/text()').extract_first()
            gamerank=gamerank.strip()
            position=mainrow.xpath('.//td[1]/div/div/span[@data-test-selector="text-horsePosition"]/text()').extract_first()
            position=position.strip()
            name=mainrow.xpath('.//td[2]/div/div/div[@class="rp-horseTable__horse"]/a/text()').extract_first()
            name=name.strip()
      
            margin_aa=mainrow.xpath('.//td[1]/div/div/span[@class="rp-horseTable__pos__length"]/span[1]/text()').extract_first()
            margin_bb=mainrow.xpath('.//td[1]/div/div/span[@class="rp-horseTable__pos__length"]/span[2]/text()').extract_first()
   
            marginvalue=0
            score=50
            if (margin_aa==None and margin_bb==None):

                margin=""
            if (margin_aa!=None and margin_bb==None):

                margin=margin_aa
            if (margin_aa!=None and margin_bb!=None):
                margin=margin_bb

            margin=margin.strip("[").strip("]")         

            margin=list(margin)
 
            marginlength = len(margin)
            if marginlength > 0:
                marginvalue = 0
                for i in range(len(margin)-1):    
                    if margin[i] == "n" or margin[i]== "h" or margin[i] == "s" or margin[i] == "d":
                        break               
                    marginvalue = marginvalue + float(margin[i])*float(math.pow(10,len(margin)-2-i))
                if margin[-1:][0] == '¼':
                    marginvalue = marginvalue + 0.25
                elif margin[-1:][0] == '¾':
                    marginvalue = marginvalue + 0.75
                elif margin[-1:][0] == '½':
                    marginvalue = marginvalue + 0.5
                elif margin[-1:][0] == 'k':              
                    marginvalue = marginvalue + 0.25   
                elif margin[-1:][0] == 's':              
                    marginvalue = marginvalue + 0.0
                    marginvalue = marginvalue_list[-1:][0] 
                elif margin[-1:][0] == 'd':            
                    if margin[0] == 's':
                        marginvalue = marginvalue + 0.0  
                        marginvalue = marginvalue_list[-1:][0] 
                    else: 
                        marginvalue = marginvalue + 0.25 
                elif margin[-1:][0] == 'e':              
                    marginvalue = marginvalue + 0.0  
                    marginvalue = marginvalue_list[-1:][0]   
                elif margin[-1:][0] == 't':              
                    marginvalue = marginvalue + 0.0
                    marginvalue = marginvalue_list[-1:][0]       
                else:
                    marginvalue=0
                    for i in range(len(margin)):  
                        marginvalue = marginvalue + float(margin[i])*float(math.pow(10,len(margin)-1-i))
                score=score+marginvalue*4  
            marginvalue_list.append(marginvalue)      
            
            jocker_allowance=mainrow.xpath('.//td[3]/div/span[@class="rp-horseTable__human__wrapper"]/sup/text()').extract_first()

            if jocker_allowance:
                pass
            else:
                jocker_allowance=""                
            age=mainrow.xpath('.//td[4]/text()').extract_first()
            age=age.strip()
            stone=mainrow.xpath('.//td[5]/span[1]/text()').extract_first()
            if stone==None:
                stone=""
            stone=stone.strip()
            pound=mainrow.xpath('.//td[5]/span[2]/text()').extract_first()
            if pound==None:
                pound=""            
            pound=pound.strip()            

            weight=stone+'-'+pound
            if stone=="" and pound=="":
                allpound=""
            else:
                allpound=int(stone)*14+int(pound)
            if gamerank=="1":
                won="1"            
            yoallowance = ""
            if int(age) == 3:
                rownumber = 8
                columnnumber = 7
                if (distance_yard>=1050 and distance_yard<=1210):
                    rownumber = 3
                elif (distance_yard>=1211 and distance_yard<=1430):
                    rownumber = 4
                elif (distance_yard>=1431 and distance_yard<=1650):
                    rownumber = 5
                elif (distance_yard>=1651 and distance_yard<=1870):
                    rownumber = 6
                elif (distance_yard>=1871 and distance_yard<=2091):
                    rownumber = 7
                elif (distance_yard>=2092 and distance_yard<=2310):
                    rownumber = 8
                if (month=="2" and int(day)<=14):
                    columnnumber = 4
                elif month == "2":
                    columnnumber = 5
                elif int(day) <= 15:
                    columnnumber = int(month)*2
                else:
                    if int(day) <= 16: 
                        columnnumber = int(month)*2
                    else:
                        columnnumber = int(month)*2+1
                
                with open('yearsold.csv') as f:
                    yocsv = csv.reader(f)
                    rowfornumber = 0
                    for row in yocsv:             
                        rowfornumber = rowfornumber+1
                        if rowfornumber == rownumber:                            
                            yoallowance = row[columnnumber]
            
            resultlist=[type_of_race,course_date,course_name,distance_yard,course_going,
                    course_class,position,name,age,weight,stone,pound,allpound,
                    jocker_allowance,yoallowance,won,score,"",marginvalue]
            if resultlist[6].isdigit():                
                self.url_result.append(resultlist)
            yield{
                'Type_of_race':type_of_race,
                'date':course_date,
                'course':course_name,
                'distance/yards':distance_yard,
                'going':course_going,
                'class':course_class,
                'position':position,
                'horse name':name,
                'age':age,
                'weight':weight,
                'stone':stone,
                'pound':pound,
                'all pound':allpound,
                'jocker_allowance':jocker_allowance,
                'allowance':yoallowance,
                'Won':won,
                'score':score,
                'margin':marginvalue      
            }
    def close(self, reason):
        age_allowance=True
        for rowaa in self.url_result:
            if int(rowaa[8])==3:
                age_allowance=False
            else:
                age_allowance=True
                break
        if age_allowance==False:
            rowcc=[]
            result_one_url=[]
            for rowbb in self.url_result:
                rowcc=rowbb
                rowcc[14]=""
                result_one_url.append(rowcc) 
        else:
            result_one_url=self.url_result           
        with open('./tmp/resultcsv.csv','a') as resultfile:
            writer = csv.writer(resultfile)   
            for row in result_one_url:             
                writer.writerow(row)
        resultfile.close()
  