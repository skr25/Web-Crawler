from __future__ import division, unicode_literals
import codecs
import re
import os
import xlrd
import requests
from urllib.request import urlopen
from urllib.request import urlretrieve
from time import sleep
from bs4 import BeautifulSoup
import openpyxl
import ssl
from collections import Counter
from urllib.request import Request
import urllib.error 
import urllib3
import xlscalls
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
USER_HEADER={'user-agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36','content-type' : 'text/html', 'accept': 'text/html','accept-charset' : 'UTF-8'}

DEFAULT_CONTEXT = ssl.create_default_context() 
DEFAULT_CONTEXT.check_hostname = False 
DEFAULT_CONTEXT.verify_mode = ssl.CERT_NONE

def checkIfValidURL(requestURL,requestProtocol):
	if requestProtocol == None:
		requestProtocol = "https"
	
	formattedURL = requestProtocol + "://" + requestURL
	request = urllib.request.Request(formattedURL, headers=USER_HEADER)
	try:
		requestData = urllib.request.urlopen(request, context=DEFAULT_CONTEXT,timeout=2)
                        

	except urllib.error.URLError as e:
		sleep(1)
		if hasattr(e, 'reason'):
			print ('We failed to reach a server.')
			return False
		elif hasattr(e, 'code'):
			print ('The server couldn\'t fulfill the request.')
			print ('Error code: ', e.code)
			return False
		else:
			return False
                        
	else:
		return True	

	return False

def getRequestData(requestURL,requestProtocol):
	if requestProtocol == None:
            	requestProtocol = "https"
	
	formattedURL = requestProtocol + "://" + requestURL
	requestData = ''
	while requestData == '':
	        try:
	                requestData = requests.get(formattedURL, verify=False,headers=USER_HEADER)
	        except:
                        sleep(1)
	return requestData

def getMetaTagsData(contentData,productName):
        custTestimonialCount = 0
        caseStudyCount = 0
        blogCount = 0
        pressCount = 0
        otherCount = 0
        phrase = []
        matchHits = 0
        productName = productName.lower()
        metaSummary = {'feature':productName,'customer-testimonial':custTestimonialCount,'case-study':caseStudyCount,'blog':blogCount,'press':pressCount,'other':otherCount}
        metaList = []
        tag_list =[]
        page = 0
        lm_date = None
        pageData = contentData.text
        soupLevel1 = BeautifulSoup(pageData, "html.parser")
        for metaTag in soupLevel1.find_all('meta'):
             #print(metaTag)
             metaAddress = metaTag.encode('ascii')
             if not b'Log In' in metaAddress:
                  try:
                      tempPageData = metaTag.get('content')
                      #print("my "+ tempPageData)
                      if tempPageData != "None":
                         content = tempPageData.lower()
                         if productName in content :
                              
                              matchHits += 1
                              #print('you got hit '+str(matchHits) + tempPageData)
                              phrase.append(tempPageData)
                  except:
                      sleep(1)
                      continue
                  soupLevel2 = BeautifulSoup(content, "html.parser")
                  sleep(1)
                  s = soupLevel2(text = re.compile(productName))
                  #print(s)
                  if s:
                       metaAddressLevel2 = metaTag.encode('ascii')
                      
                       if b'customer-testimonial' in metaAddressLevel2:
                            metaSummary['customer-testimonial'] += 1
                       elif b'case-study' in metaAddressLevel2:
                            metaSummary['case-study'] += 1
                       elif b'blog' in metaAddressLevel2:
                            metaSummary['blog'] += 1
                       elif b'press' in metaAddressLevel2:
                            metaSummary['press'] += 1
                       else:
                            metaSummary['other'] += 1
                       for tag in s:
                            parentHtml = tag.parent.name
                            tag_list.append(parentHtml)
                       page += 1
                       metaInfoSummary = {'feature':productName,'tagName':'meta','phrase':s,'address':metaAddressLevel2,'counts':Counter(tag_list),'addressUrl':metaTag.get('href'),'urlAdd':lm_date}
                       metaList.append(metaInfoSummary)                      
        return metaSummary,metaList,phrase,matchHits

def getHrefTagsData(contentData,productName,tagName,pageURL):
        custTestimonialCount = 0
        caseStudyCount = 0
        blogCount = 0
        pressCount = 0
        otherCount = 0
        phrase = []
        matchHits = 0
        productName = productName.lower()
        hrefSummary = {'feature':productName,'customer-testimonial':custTestimonialCount,'case-study':caseStudyCount,'blog':blogCount,'press':pressCount,'other':otherCount}
        hrefList = []
        tag_list =[]
        page = 0
        lm_date = None
        pageData = contentData.text

        soupLevel1 = BeautifulSoup(pageData, "html.parser")
        for hrefTag in soupLevel1.find_all(tagName):
             #print(metaTag)
             hrefAddress = hrefTag.encode('ascii')
             if not b'Log In' in hrefAddress:
                  try:
                      #print(hrefTag.get('href'))
                      hrefURL = hrefTag.get('href')
                      if hrefURL[0] == '/':
                            hrefAbsURL = pageURL + hrefURL
                      else:
                            hrefAbsURL = hrefURL
                      #print('before request %s'  % hrefAbsURL)
                      hrefRequest = requests.get(hrefAbsURL, verify=False, headers=USER_HEADER,timeout=3)
                      #print('after request %s'  % hrefAbsURL)
                      tempPageData = hrefRequest.text
                      #print(tempPageData)
                      #gg=input("enter a key")
                      
                      if tempPageData != "None":
                         content = tempPageData.lower()
                         #print('there is content')
                         lm_date = hrefRequest.headers.get('Last-Modified','None')
                         if productName in content :
                              
                              matchHits += 1
                              #print(matchHits)
                      #else:
                          #print('no content. It is None.')  
                              
                  except:
                      sleep(1)
                      #print('not a valid url')
                      continue
                  #print('in soup')
                  soupLevel2 = BeautifulSoup(content, "html.parser")
                  sleep(1)
                  #print(productName)
                  s = soupLevel2(text = re.compile(productName))
                  #sleep(1)
                  #print(s)
                  if s:
                       phrase.append(s)
                      
                       hrefAddressLevel2 = hrefTag.encode('ascii')
                       if b'customer-testimonial' in hrefAddressLevel2:
                            hrefSummary['customer-testimonial'] += 1
                       elif b'case-study' in hrefAddressLevel2:
                            hrefSummary['case-study'] += 1
                       elif b'blog' in hrefAddressLevel2:
                            hrefSummary['blog'] += 1
                       elif b'press' in hrefAddressLevel2:
                            hrefSummary['press'] += 1
                       else:
                            hrefSummary['other'] += 1
                       for tag in s:
                            parentHtml = tag.parent.name
                            tag_list.append(parentHtml)
                       page += 1
                       hrefInfoSummary = {'feature':productName,'tagName':tagName,'phrase':s,'address':hrefAddressLevel2,'counts':Counter(tag_list),'addressUrl':hrefTag.get('href'),'urlAdd':lm_date}
                       hrefList.append(hrefInfoSummary)                      
        return hrefSummary,hrefList,phrase,matchHits


def consolidateSummary(summaryList1,summaryList2,summaryList3):
     resultSummary= {'feature':None,'customer-testimonial':None,'case-study':None,'blog':None,'press':None,'total-posts':None}
     resultSummary['customer-testimonial'] = summaryList1['customer-testimonial']+summaryList2['customer-testimonial']+summaryList3['customer-testimonial']
     resultSummary['case-study'] = summaryList1['case-study']+summaryList2['case-study']+summaryList3['case-study']
     resultSummary['blog'] = summaryList1['blog']+summaryList2['blog']+summaryList3['blog']
     resultSummary['press'] = summaryList1['press']+summaryList2['press']+summaryList3['press']
     resultSummary['total-posts'] = summaryList1['other']+summaryList2['other']+summaryList3['other']
     resultSummary['feature'] = summaryList1['feature']
     return resultSummary

if __name__ == "__main__":

    linkList = xlscalls.getLinkListFromWB('/share/Public/crawlerSKR/linklist_1.xlsx')
    productList = xlscalls.getProductListFromWB('/share/Public/crawlerSKR/product.xlsx',5,3)
    print(productList)
    print(linkList)

    for eachLink in linkList:
        userXLList = []
        miXLList = []
        for eachProduct in productList:
           print(eachProduct)
    
           if checkIfValidURL(eachLink,'https') == True:
               pageData = getRequestData(eachLink,'https')
               pageURL = 'https://'+eachLink
               print("Inside Https")
           elif checkIfValidURL(eachLink,'http') == True:
               pageData = getRequestData(eachLink,'http')
               pageURL = 'http://'+ eachLink
               print("Inside Http")
           else:
               print('website does not exists')
               exit()
    
           summ,summlist,phrase,hits  = getMetaTagsData(pageData,eachProduct)
           #print(summ)
           #print(len(phrase))
           #print(phrase)
           #print(summlist)
           #print(phrase)
           #print(hits)
           hrefsumm,hrefsummlist,hrefphrase,hrefhits  = getHrefTagsData(pageData,eachProduct,'a',pageURL)  
           #print(hrefsumm)
           #print(len(hrefphrase))
           #print(hrefphrase)
           #print(hrefsummlist)
           #print(hrefphrase)
           #print(hrefhits)
           linksumm,linksummlist,linkphrase,linkhits  = getHrefTagsData(pageData,eachProduct,'link',pageURL)  
           #print(linksumm)
           #print(len(linkphrase))
           #print(linkphrase)
       
           resultList = consolidateSummary(summ,hrefsumm,linksumm)

           userXLList.append(resultList)
           miXLList.extend(summlist)
           miXLList.extend(hrefsummlist)
           miXLList.extend(linksummlist)
           #input('press a key')
           print(userXLList)
       
        xlscalls.createUserWB(userXLList,miXLList,eachLink)
        #xlscalls.createUserDebugWB(userXLList,miXLList,eachLink)

