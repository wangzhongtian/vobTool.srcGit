#-*- coding: UTF-8
from __future__ import print_function
import glob
template1="""
class Zt(object):\n
        test=0
        def __init__(self):
        	pass
        def __getattribute__(self, name) :
            try:
                print();print();print()
                excfile=eval('Zt.'+name,globals())
                print(excfile)
                execfile( excfile,globals())
                print( "tt."+name)
            except Exception as r:
                print(r,type(r) )
                pass
            finally:
               print();print();print()
               sys.exit()

            return 0
	"""
template2="""
class Zt(object):\n
        test=0
        def __init__(self):
        	pass
        def __getattribute__(self, name) :
                print();print();print()
                excfile=eval('Zt.'+name,globals())
                print(excfile)
                execfile( excfile,globals())
                print( "tt."+name)
                print();print();print()
                sys.exit()
                return 0
	"""
template =template2
methodTpl="""
        @classmethod
        def tt{0:}(cls):
            execfile( '{0:}.py')
            return 0 \n
"""
propertyTpl="""
        tt{0:}='{1:}'
"""
import os
import sys
class test():
	# @classmethod
	def __init__(self):
		libpath=os.environ[ "libpath"].split(";")
		# print(libpath)
		files=[]
		for libp in libpath:
			fileters= libp+"\\*.py"
			# print(libp,fileters)
			files += glob.glob(fileters)
		files += glob.glob("*.py")
		cmdstr =template
		for f in files:
			f1=f[:-3]
			path ,file = os.path.split( f1) 
			cmdstr += propertyTpl.format( file , f)
			# cmdstr += "execfile('{}')".format(f1 )
		# print( cmdstr) 
		print(os.path.abspath("."))
		exec cmdstr  in  globals() 

def main():

	# execfile("Shell.py")
	temobj=test()
	globals()["tt"] = Zt()
	print(" Shell ready...................")
main()
主机名号="易县004号"
主机名号="易县003号"
主机名号="易县006号"
主机名号="西黑山001号"

print("主机名号={}".format(主机名号))

# execfile("e:\ipy\libsrc\shell.py")
# main()
##  ## 将搜索LIbpath 以及当前文件加下定义的所有Py文件并生成一个快捷方式
##  ## 这个快捷方式将可以通过如下的方式方便调用
# tt.ttcompileIPy2DLL