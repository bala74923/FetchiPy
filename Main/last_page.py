#prerequiste 1)contest name 2) requests
import requests

#CONTEST_NAME = "weekly contest 319"
#contest_name = CONTEST_NAME.lower().strip().replace(' ','-')


def get_last_page_of_contest(contest_name):
    start = 1
    end = 2048
    last_page = 1
    # find last occuring leetcode rank page
    while start<=end:
        try:
            page = int((start+end)/2)
            API_URL_FMT = 'https://leetcode.com/contest/api/ranking/{}/?pagination={}&region=global'
            page_link = f'https://leetcode.com/contest/{contest_name}/ranking/{page}/'
            url = API_URL_FMT.format(contest_name, page)
            print(url)
            resp = requests.get(url).json()
            print(resp)
            objs = resp['total_rank']
            #fact we know len(objs)==0 then not valid
            if len(objs)>0:#may or may not be the last page
                last_page = page
                start = page+1
            else:
                end=  page-1
        except Exception as e:
            print('error for page',page)
            end = page-1
    print(last_page," is the last page")
    return last_page