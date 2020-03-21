
import clr
import sys
import System
import System.IO

def getFullNames1(dllname):
  filename=None
  filename = "./"+dllname
  if (   System.IO.File.Exists(filename) ):
      print("Load Dlls:" ,filename)
      return dllname
  for p in System.Environment.GetEnvironmentVariable("libpath").split(";"):
    filename = p+"/"+dllname
    if (   System.IO.File.Exists(filename) ):
      print("Load Dlls:" ,filename)
      return filename

def getFullNames(dllname):
    filename=None
    filename = "./"+dllname
    if ( System.IO.File.Exists(filename) ):
        print("Load Dlls:" ,filename)
        return dllname
    for p in sys.path:
        print(p)
        filename = p+"/"+dllname
        if (   System.IO.File.Exists(filename) ):
            print("Load Dlls:" ,filename)
            return filename