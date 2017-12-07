import glob
import os
from os import listdir
__author__ = 'maor'
# filename = ''
# minvoters = 99999999
# for file in listdir('3Cand'):
#      tmpfile = open('3Cand\\' + file, "r")
#      lines = tmpfile.readlines()
#      tmpnum = int(lines[4].split(',')[0])
#      if tmpnum < minvoters:
#          minvoters = tmpnum
#          filename = tmpfile.name
#      tmpfile.close()
#
# print str(minvoters) + " " + filename
#
#


for fname in glob.glob('SushiandOther/3CandReduct/Original/*.txt'):
    tmpfile = open(fname, "r")
    NumberOfVoters = 9
    summingVotes = 0
    #text_file = open("ED-00004-00000001.soc", 'r')
    lines = tmpfile.readlines()
    reductfile = open('/'.join(os.path.dirname(fname).split('/')[:-1]) + '/' + str(NumberOfVoters) + 'Voters/' + lines[0].strip() + "Cand_" + str(NumberOfVoters) + "Voters-" + tmpfile.name.split('_')[1], "w")
    count = 0
    totalVotes = 0
    for line in lines:
        if count < int(lines[0]) + 1:
            reductfile.write(line)
        else:
            splittedline = line.split(',')
            if count == (int(lines[0]) + 1):
                totalVotes = int(splittedline[0])
                reductfile.write(str(NumberOfVoters) + "," + str(NumberOfVoters) + "," + splittedline[2])
            else:
                current = int(round((float(splittedline[0]) / totalVotes) * NumberOfVoters))
                summingVotes += current
                reductfile.write(str(current) + line[len(splittedline[0]):])
        count += 1

    tmpfile.close()
    reductfile.close()
    if not summingVotes == NumberOfVoters:
        os.rename(reductfile.name, os.path.dirname(reductfile.name) + '/Not Complete/NC-' + os.path.basename(reductfile.name) )