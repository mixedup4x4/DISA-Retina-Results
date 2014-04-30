import subprocess
import time
import sys
import re

class checkIfUp:
    __shellPings = []
    __shell2Nbst = []
    __ipsToCheck = []
    checkedIps = 0
    onlineIps = 0
    unreachable = 0
    timedOut = 0
    upIpsAddress = []
    computerName = []
    completeMacAddress = []
    executionTime = 0
    
    def __init__(self,fromIp,toIp):
        startTime = time.time()
        self.fromIp = fromIp # from 192.168.1.x
        self.toIp = toIp # to 192.168.x.x
        self.__checkIfIpIsValid(fromIp)
        self.__checkIfIpIsValid(toIp)
        self.__getRange(fromIp,toIp)
        self.__shellToQueue()
        #self.__checkIfUp() # run by the shellToQueue queue organizer
        self.__computerInfoInQueue()
        endTime = time.time()
        self.executionTime = round(endTime - startTime,3)
        
    def __checkIfIpIsValid(self,ip):
        def validateRange(val):
            # valid range => 1 <-> 255
            try:
                val = int(val)
                if val < 0 or val > 255:
                    print "Invalid IP Range ("+str(val)+")"
                    sys.exit(0)
            except:
                print "Invalid IP"
                sys.exit(0)
        ip = ip.split(".")
        firstVal = validateRange(ip[0])
        secondVal = validateRange(ip[1])
        thirdVal = validateRange(ip[2])
        fourthVal = validateRange(ip[3])
        return True
    
    def __getRange(self,fromIp,toIp):
        fromIp = fromIp.split(".")
        toIp = toIp.split(".")

        # toIp must be > fromIp
        def ip3chars(ipBlock):
            # input 1; output 001
            ipBlock = str(ipBlock)
            while len(ipBlock) != 3:
                ipBlock = "0"+ipBlock
            return ipBlock
        fromIpRaw = ip3chars(fromIp[0])+ip3chars(fromIp[1])+ip3chars(fromIp[2])+ip3chars(fromIp[3])
        toIpRaw = ip3chars(toIp[0])+ip3chars(toIp[1])+ip3chars(toIp[2])+ip3chars(toIp[3])
        if fromIpRaw > toIpRaw:
            # if from is bigger switch the order
            temp = fromIp
            fromIp = toIp
            toIp = temp

        currentIp = [0,0,0,0]
        # all to integers
        currentIp0 = int(fromIp[0])
        currentIp1 = int(fromIp[1])
        currentIp2 = int(fromIp[2])
        currentIp3 = int(fromIp[3])
        toIp0 = int(toIp[0])
        toIp1 = int(toIp[1])
        toIp2 = int(toIp[2])
        toIp3 = int(toIp[3])

        firstIp = str(currentIp0)+"."+str(currentIp1)+"."+str(currentIp2)+"."+str(currentIp3)
        self.__ipsToCheck = [firstIp]
        while currentIp3 != toIp3 or currentIp2 != toIp2 or currentIp1 != toIp1 or currentIp0 != toIp0:
            currentIp3 += 1
            if currentIp3 > 255:
                currentIp3 = 0
                currentIp2 += 1
                if currentIp2 > 255:
                    currentIp2 = 0
                    currentIp1 += 1
                    if currentIp1 > 255:
                        currentIp1 = 0
                        currentIp0 += 1
            addIp = str(currentIp0)+"."+str(currentIp1)+"."+str(currentIp2)+"."+str(currentIp3)
            self.__ipsToCheck.append(addIp)
        
    def __shellToQueue(self):
        # write them in the shell queue
        maxPingsAtOnce = 200
        currentQueuedPings = 0
        for pingIp in self.__ipsToCheck:
            proc = subprocess.Popen(['ping','-n','1',pingIp],stdout=subprocess.PIPE,shell=True)
            self.__shellPings.append(proc)
            currentQueuedPings += 1
            if currentQueuedPings >= maxPingsAtOnce:
                #execute shells
                self.__checkIfUp()
                currentQueuedPings = 0
                self.__shellPings = []
        self.__checkIfUp() # execute last queue
            
    def __checkIfUp(self):
        # execute the shells & determine whether the host is up or not
        for shellInQueue in self.__shellPings:
            pingResult = ""
            shellInQueue.wait()
            while True:
                line = shellInQueue.stdout.readline()
                if line != "":
                    pingResult += line
                else:
                    break;
            self.checkedIps += 1
            if 'unreachable' in pingResult:
                self.unreachable += 1
            elif 'timed out' in pingResult:
                self.timedOut += 1
            else:
                self.onlineIps += 1
                currentIp = self.__ipsToCheck[self.checkedIps-1]
                self.upIpsAddress.append(currentIp)

    def __computerInfoInQueue(self):
        # shell queue for online hosts
        maxShellsAtOnce = 255
        currentQueuedNbst = 0
        for onlineIp in self.upIpsAddress:
            proc = subprocess.Popen(['\\Windows\\sysnative\\nbtstat.exe','-a',onlineIp],stdout=subprocess.PIPE,shell=True)
            self.__shell2Nbst.append(proc)
            currentQueuedNbst += 1
            if currentQueuedNbst >= maxShellsAtOnce:
                # execute shells
                self.__gatherComputerInfo()
                currentQueuedNbst = 0
                self.__shell2Nbst = []
        self.__gatherComputerInfo() # execute last queue

    def __gatherComputerInfo(self):
        # execute the shells and find host Name and MAC
        for shellInQueue in self.__shell2Nbst:
            nbstResult = ""
            shellInQueue.wait()

            computerNameLine = ""
            macAddressLine = ""
            computerName = ""
            macAddress = ""
            while True:
                line = shellInQueue.stdout.readline()
                if line != "":
                    if '<00>' in line and 'UNIQUE' in line:
                        computerNameLine = line
                    if 'MAC Address' in line:
                        macAddressLine = line
                else:
                    break;
                
            computerName = re.findall('([ ]+)(.*?)([ ]+)<00>', computerNameLine)
            macAddress = re.findall('([A-Z0-9]+)-([A-Z0-9]+)-([A-Z0-9]+)-([A-Z0-9]+)-([A-Z0-9]+)-([A-Z0-9]+)',macAddressLine)
            try:
                self.computerName.append(computerName[0][1])
            except:
                self.computerName.append("")

            completeMacAddress = ""
            firstMacElement = 0
            try:
                for macEach in macAddress[0]:
                    if firstMacElement == 0:
                        firstMacElement += 1
                    else:
                        completeMacAddress += ":"
                    completeMacAddress += macEach
                firstMacElement = 0
            except:
                completeMacAddress = ""
            self.completeMacAddress.append(completeMacAddress)
                        
    def readValue(self):
        # debugging use only
        ips = []
        for ip in self.completeMacAddress:
            ips.append(ip)
        return ips

print "\t\t---LANScanner v1.0---\n"
# brief tutorial
print "Sample input data:"
print "FromIP: 192.168.1.50"
print "ToIP: 192.168.1.20"
print "---"
# input
fromIp = raw_input("From: ")
toIp = raw_input("To: ")
# enter values to class
userRange = checkIfUp(fromIp,toIp)
# read class values
print ""
#print userRange.readValue() # debugging use only
print "Checked",userRange.checkedIps,"IPs"
print ""
print "Online:",str(userRange.onlineIps)+"/"+str(userRange.checkedIps)
print "Unreachable:",userRange.unreachable,"Timed out:",userRange.timedOut
print "" # newline
print "Online IPs:"
print "IP\t\tNAME\t\tMAC"
counter = 0
for onlineIp in userRange.upIpsAddress:
    print onlineIp+"\t"+userRange.computerName[counter]+"\t"+userRange.completeMacAddress[counter]
    counter += 1
print ""
print "Took",userRange.executionTime,"seconds"
