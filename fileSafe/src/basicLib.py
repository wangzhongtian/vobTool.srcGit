#-*- coding: UTF8
from __future__ import print_function
import sys
def openfile(filename, mode="wt",encoding="UTF8"):
    if "2.7"  in sys.version:
        import codecs
        curfile=codecs.open(filename ,mode ,encoding)
        # curfile= open( tgrFileName ,"wt" ,encoding="UTF8")
    else:
        curfile= open( filename ,mode ,encoding)
    return curfile


        