import glob
import os
from os import listdir
import itertools
import operator

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
Cands = 3
MAXCOMB = 10

for file in glob.glob('4Cand/Original/*.soc'):
    tmpfile = open(file, "r")
    #text_file = open("ED-00004-00000001.soc", 'r')
    lines = tmpfile.readlines()

    candinfile = int(lines[0])
    for idx, comb in enumerate(itertools.combinations(range(1, candinfile + 1), 3)):
        if idx == MAXCOMB:
            break
        fileperm = {}
        for perm in itertools.permutations(comb):
            fileperm[tuple([str(x) for x in perm])] = 0
        reductfile = open('4to3Cand/' + str(Cands) + "Cand_" + '-'.join(tmpfile.name.split('-')[1:]).split('.')[
            0] + "-" + str(idx + 1) + ".txt", "w")
        reductfile.write(str(Cands) + '\n')
        for count, line in enumerate(lines[1:]):
            if comb.__contains__(count + 1):
                reductfile.write(line)
            elif count + 1 == candinfile:
                break
        lines2 = [x.strip().split(',') for x in lines]
        reductfile.write(','.join(lines2[candinfile + 1][:-1]) + ',' + str(len(fileperm)) + '\n')
        for line in lines2[candinfile + 2:]:
            tmp = line[1:]
            for x in range(1, candinfile + 1):
                if not comb.__contains__(x):
                    tmp.remove(str(x))
            fileperm[tuple(tmp)] += int(line[0])

        sorted_x = sorted(fileperm.iteritems(), key=operator.itemgetter(1))
        sorted_x.reverse()
        for ww in sorted_x:
            reductfile.write(str(ww[1]) + ',' + ','.join(ww[0]) + '\n')
        reductfile.close()