# -*- coding: utf-8 -*-

import requests
import time
import urllib as ub
import xlrd
import json
from bs4 import BeautifulSoup
import execjs #必须，需要先用pip 安装，用来执行js脚本

import os
import sys
BASEDIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(BASEDIR)

print(BASEDIR)
class Py4Js():     
  def __init__(self):  
    self.ctx = execjs.compile(""" 
    function TL(a) { 
    var k = ""; 
    var b = 406644; 
    var b1 = 3293161072;       
    var jd = "."; 
    var $b = "+-a^+6"; 
    var Zb = "+-3^+b+-f";    
    for (var e = [], f = 0, g = 0; g < a.length; g++) { 
        var m = a.charCodeAt(g); 
        128 > m ? e[f++] = m : (2048 > m ? e[f++] = m >> 6 | 192 : (55296 == (m & 64512) && g + 1 < a.length && 56320 == (a.charCodeAt(g + 1) & 64512) ? (m = 65536 + ((m & 1023) << 10) + (a.charCodeAt(++g) & 1023), 
        e[f++] = m >> 18 | 240, 
        e[f++] = m >> 12 & 63 | 128) : e[f++] = m >> 12 | 224, 
        e[f++] = m >> 6 & 63 | 128), 
        e[f++] = m & 63 | 128) 
    } 
    a = b; 
    for (f = 0; f < e.length; f++) a += e[f], 
    a = RL(a, $b); 
    a = RL(a, Zb); 
    a ^= b1 || 0; 
    0 > a && (a = (a & 2147483647) + 2147483648); 
    a %= 1E6; 
    return a.toString() + jd + (a ^ b) 
  };      
  function RL(a, b) { 
    var t = "a"; 
    var Yb = "+"; 
    for (var c = 0; c < b.length - 2; c += 3) { 
        var d = b.charAt(c + 2), 
        d = d >= t ? d.charCodeAt(0) - 87 : Number(d), 
        d = b.charAt(c + 1) == Yb ? a >>> d: a << d; 
        a = b.charAt(c) == Yb ? a + d & 4294967295 : a ^ d 
    } 
    return a 
  } 
 """)            
  def getTk(self,text):  
      return self.ctx.call("TL",text)
def buildUrl(text,tk):
  baseUrl='https://translate.google.cn/translate_a/single'
  baseUrl+='?client=t&'
  baseUrl+='s1=auto&'
  baseUrl+='t1=zh-CN&'
  baseUrl+='h1=zh-CN&'
  baseUrl+='dt=at&'
  baseUrl+='dt=bd&'
  baseUrl+='dt=ex&'
  baseUrl+='dt=ld&'
  baseUrl+='dt=md&'
  baseUrl+='dt=qca&'
  baseUrl+='dt=rw&'
  baseUrl+='dt=rm&'
  baseUrl+='dt=ss&'
  baseUrl+='dt=t&'
  baseUrl+='ie=UTF-8&'
  baseUrl+='oe=UTF-8&'
  baseUrl+='otf=1&'
  baseUrl+='pc=1&'
  baseUrl+='ssel=0&'
  baseUrl+='tsel=0&'
  baseUrl+='kc=2&'
  baseUrl+='tk='+str(tk)+'&'
  baseUrl+='q='+text
  return baseUrl


def get_excel_data(file_url):
    workbook = xlrd.open_workbook(file_url) # [u'sheet1', u'sheet2']
    # sheet0 = workbook.sheet_names()[0]
    table = workbook.sheet_by_index(0) # sheet索引从0开始
    #sheet2 = workbook.sheet_by_name('sheet2')

    # sheet的名称，行数，列数
    # print sheet2.name,sheet2.nrows,sheet2.ncols

    # 命令所属列数
    key_col_index = 1
    # 翻译文本所属的列数
    content_col_index = 2
    # 翻译文本英语递增数（相对key起始行）
    en_step = 1
    # 翻译文本法语递增数（相对key起始行）
    fr_step = 2
    # 返回结果
    result = {}
    # 其实行
    count = 1
    # 保存文件的根路径
    save_file_base_path = os.path.join(BASEDIR, 'files')
    if not os.path.exists(save_file_base_path):
        os.mkdir(save_file_base_path)
    max_rows = table.nrows
    while count < max_rows:
        key_str = str(table.cell(count, key_col_index).value)
        file_base_path = os.path.join(save_file_base_path, key_str)
        print(file_base_path)
        if not os.path.exists(file_base_path):
            os.mkdir(file_base_path)
        
        file_gender_path = os.path.join(file_base_path, 'female')
        if not os.path.exists(file_gender_path):
            os.mkdir(file_gender_path)

        for lang in ['en', 'fr']:
            file_language_path = os.path.join(file_gender_path, lang)
            if not os.path.exists(file_language_path):
                os.mkdir(file_language_path)

        en_content = table.cell(count + en_step, content_col_index).value.encode('utf8')
        fr_content = table.cell(count + fr_step, content_col_index).value.encode('utf8')
        result[key_str] = {
            'en': (en_content.strip().split(':')[1], os.path.join(file_gender_path, 'en')),
            'fr': (fr_content.strip().split(':')[1], os.path.join(file_gender_path, 'fr'))
        }
        count += 3

    return result

def translate(data, force_down=False):
    js=Py4Js()
    header={
    'authority':'translate.google.cn',
    'method':'GET',
    'path':'',
    'scheme':'https',
    'accept':'*/*',
    'accept-encoding':'gzip, deflate, br',
    'accept-language':'zh-CN,zh;q=0.9',
    'cookie':'',
    'user-agent':'Mozilla/5.0 (Windows NT 10.0; WOW64)  AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.108 Safari/537.36',
    'x-client-data':'CIa2yQEIpbbJAQjBtskBCPqcygEIqZ3KAQioo8oBGJGjygE='
    }

    for file_name, _data in data.items():
        time.sleep(1)
        for lang, (text, base_file_path) in _data.items():
            file_path = os.path.join(base_file_path, file_name)
            tk = js.getTk(text)
            host_url = 'https://translate.google.cn/translate_tts?'

            params = {
                'ie': 'UTF-8',
                'q': text,
                'tl': lang,
                'total': 1,
                'idx':0,
                'textlen':len(text),
                'client':'webapp',
                'tk':tk
            }
            data_url = ub.urlencode(params)
            file_abspath = file_path + ".wav"
            if os.path.exists(file_abspath) and not force_down:
                continue
            try:

                newfile=open(file_path+".wav","wb")
                context = requests.get(host_url + data_url, timeout = 3000)
                for data in context.iter_content(chunk_size=1024):
                    if data:
                        newfile.write(data)
                newfile.close()
                print(file_path)
            except Exception as e:
                print('tk: {}'.format(tk))
                print('file_abspath: {}'.format(file_abspath))
                print(e)

def run():
    print('进来了')
    res = get_excel_data(os.path.join(BASEDIR, 'static/中文-英语-法语.xls'))
    print(res)
    translate(res)