import os
import subprocess
import sys


def runDissect(inputString):
    print("Hey")
    outputFile = "output.json"
    FNULL = open(os.devnull, 'w')
    print("here")
    args = "r6Dissect.exe" + inputString + "-x" + outputFile
    subprocess.call(args,stdout=FNULL,stderr=FNULL,shell=True)
    print('worked')

runDissect("Match-2023-07-12_21-50-26-85")