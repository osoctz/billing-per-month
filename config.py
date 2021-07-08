import configparser
import os

conf = configparser.ConfigParser()
conf.read(os.curdir+'/config.ini', encoding='utf-8')

def get_base(key):

    return conf.get('base',key)

def get_ext(key):
    return conf.get('ext', key)
#
#
# if __name__ == '__main__':
#
#     print(get_base('company'))