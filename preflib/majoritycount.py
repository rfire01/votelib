import glob
import os
import itertools
import sys

import xlsxwriter


__author__ = 'maor'
nCands = 3
nVoters = 3
#foldCands = str(nCands) + 'Cand/' + str(nVoters) + 'Voters/*.txt
#foldCands = str(nCands) + 'Cand/Original/*.soc'
#foldCands = '4to3Cand/' + str(nVoters) + 'Voters/*.txt'
foldCands = 'SushiandOther/3CandReduct/' + str(nVoters) + 'Voters/*.txt'

class prefLib:
    def __init__(self, prefFile, filename):
        prefFile = [x.strip().split(',') for x in prefFile]
        self.filen = filename
        self.numCands = int(prefFile[0][0])
        self.nameC1 = prefFile[1][1]
        self.nameC2 = prefFile[2][1]
        self.nameC3 = prefFile[3][1]
        self.CandsTuple = (str(prefFile[1][0]), str(prefFile[2][0]), str(prefFile[3][0]))
        prefstart = 5
        if self.numCands == 4:
            prefstart = 6
            self.nameC4 = prefFile[4][1]
            self.CandsTuple += (str(preFile[4][0]), )
        self.totalVotes = int(prefFile[prefstart - 1][0])
        self.PrefList = list(prefFile[prefstart:])
        self.vCand4, self.vCand3, self.vCand2, self.vCand1 = 0, 0, 0, 0

    def votes(self, votefor, votes):
        candid = str(self.CandsTuple.index(votefor)+1)
        if candid == '1':
            self.vCand1 += votes
        elif candid == '2':
            self.vCand2 += votes
        elif candid == '3':
            self.vCand3 += votes
        elif candid == '4':
            self.vCand4 += votes

    def PluralityCount(self):
        self.vCand4, self.vCand3, self.vCand2, self.vCand1 = 0, 0, 0, 0
        for pref in self.PrefList:
            self.votes(pref[1], int(pref[0]))

    def BordaCount(self):
        self.vCand4, self.vCand3, self.vCand2, self.vCand1 = 0, 0, 0, 0
        for pref in self.PrefList:
            self.votes(pref[1], int(pref[0]) * (self.numCands - 1))
            self.votes(pref[2], int(pref[0]) * (self.numCands - 2))
            if self.numCands == 4:
                self.votes(pref[3], int(pref[0] * (self.numCands - 3)))

    def Winner(self, minmax=False):
        sum = {'1': self.vCand1, '2': self.vCand2, '3': self.vCand3}
        if self.numCands == 4:
            sum['4'] = self.vCand4
        maxval = max(sum.values())
        if minmax:
            maxval = min(sum.values())
        first = ''
        for x in sum:
            if sum[x] == maxval:
                first += x + ', '
        return first[:-2], maxval

    def Copeland(self):
        self.vCand4, self.vCand3, self.vCand2, self.vCand1 = 0, 0, 0, 0
        cands = self.CandsTuple
        """
        cands = range(1, self.numCands + 1)
        cands = [str(x) for x in cands]
        """
        for cand1, cand2 in itertools.combinations(cands, 2):
            paircands = {cand1: 0, cand2: 0}
            for pref in self.PrefList:
                if pref[1:].index(cand1) < pref[1:].index(cand2):
                    paircands[cand1] += int(pref[0])
                else:
                    paircands[cand2] += int(pref[0])
            if paircands[cand1] > paircands[cand2]:
                self.votes(cand1, 1)
            elif paircands[cand2] > paircands[cand1]:
                self.votes(cand2, 1)
            else:
                self.votes(cand1, 0.5)
                self.votes(cand2, 0.5)

    def Minimax(self):
        self.vCand4, self.vCand3, self.vCand2, self.vCand1 = 0, 0, 0, 0
        cands = self.CandsTuple
        """
        cands = range(1, self.numCands + 1)
        cands = [str(x) for x in cands]
        """
        table = {}
        self.minmaxR = {}
        for cand1, cand2 in itertools.combinations(cands, 2):
            if not table.has_key(cand1):
                table[cand1] = {}
            if not table.has_key(cand2):
                table[cand2] = {}
            paircands = {cand1: 0, cand2: 0}
            for pref in self.PrefList:
                if pref[1:].index(cand1) < pref[1:].index(cand2):
                    paircands[cand1] += int(pref[0])
                else:
                    paircands[cand2] += int(pref[0])
            table[cand1][cand2] = [paircands[cand1], paircands[cand2]]
            table[cand2][cand1] = [paircands[cand2], paircands[cand1]]
        for cand in table:
            maxlostop, maxmargins, maxopp = 0, -sys.maxint, 0
            for opp in table[cand]:
                ca, op = table[cand][opp]
                if ca < op:
                    maxlostop = maxlostop if maxlostop > op else op
                maxmargins = maxmargins if (op - ca) < maxmargins else (op - ca)
                maxopp = maxopp if maxopp > op else op
            self.minmaxR[cand] = {'WPO': maxopp, 'WPD': maxlostop, 'WPDM': maxmargins}

    def MinMaxMetric(self, matrix):
        self.vCand4, self.vCand3, self.vCand2, self.vCand1 = 0, 0, 0, 0
        for cand in self.minmaxR:
            self.votes(cand, self.minmaxR[cand][matrix])

    def toTuple(self, minmax=False):
        if self.numCands == 4:
            return (self.vCand1, self.vCand2, self.vCand3, self.vCand4, self.Winner(minmax)[0], self.totalVotes)
        else:
            return (self.vCand1, self.vCand2, self.vCand3, self.Winner(minmax)[0], self.totalVotes)


workbook = xlsxwriter.Workbook(foldCands[:-5] + 'Results.xlsx')
plurformat = workbook.add_format({'font_color': 'red'})
bordformat = workbook.add_format({'font_color': 'green'})
copeformat = workbook.add_format({'font_color': 'blue'})

minmaxXoYform = workbook.add_format({'font_color': '#B45F04'})
minmaxXoYLform = workbook.add_format({'font_color': 'orange'})
minmaxMarginsform = workbook.add_format({'font_color': '#0489B1'})

mergeformat = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

candsRow3 = (
    'Candidate 1', 'Candidate 2', 'Candidate 3', 'Winner', 'Total Votes', 'Has Condorcet?', '', 'Filename')
candsRow4 = (
    'Candidate 1', 'Candidate 2', 'Candidate 3', 'Candidate 4', 'Winner', 'Total Votes', 'Has Condorcet?', '',
    'Filename')
candsTitle = candsRow3 if nCands == 3 else candsRow4
pluralitysheet = workbook.add_worksheet("Plurality")
bordasheet = workbook.add_worksheet("Borda")
copelandsheet = workbook.add_worksheet("Copeland")
minmaxXoY = workbook.add_worksheet("minmax X above Y")
minmaxLXoY = workbook.add_worksheet("minmax lost X above Y")
minmaxMargins = workbook.add_worksheet("minmax Margins")
minmaxmetricComp = workbook.add_worksheet("minmax Comparison")
comparesheet = workbook.add_worksheet("Comparison")

pluralitysheet.write_row('A1', candsTitle)
bordasheet.write_row('A1', candsTitle)
copelandsheet.write_row('A1', candsTitle)
minmaxXoY.write_row('A1', candsTitle)
minmaxLXoY.write_row('A1', candsTitle)
minmaxMargins.write_row('A1', candsTitle)

comparesheet.write_row('A1', candsTitle[:-3] + ('Method', 'Has Condorcet?', 'Filename'))
minmaxmetricComp.write_row('A1', candsTitle[:-3] + ('Method', 'Has Condorcet?', 'Filename'))

tmp = 0
tmp2 = 0
for idx, fval in enumerate(glob.glob(foldCands)):
    with open(fval, 'r') as ff:
        filecontent = ff.readlines()
    preFile = prefLib(filecontent, fval)

    preFile.Copeland()
    copetup = preFile.toTuple()
    cWin, chkCondorcet = preFile.Winner()
    isCondorcet = True if chkCondorcet == (preFile.numCands - 1) else False
    copelandsheet.write_row(idx + 1, 0, copetup + (isCondorcet, '', os.path.basename(fval)))

    preFile.Minimax()
    preFile.MinMaxMetric('WPO')  #'WPD' 'WPDM'
    wpotup = preFile.toTuple(minmax=True)
    wpoWin = preFile.Winner(minmax=True)[0]
    minmaxXoY.write_row(idx + 1, 0, wpotup + (isCondorcet, '', os.path.basename(fval)))

    preFile.MinMaxMetric('WPD')  #'WPO' 'WPDM'
    wpdtup = preFile.toTuple(minmax=True)
    wpdWin = preFile.Winner(minmax=True)[0]
    minmaxLXoY.write_row(idx + 1, 0, wpdtup + (isCondorcet, '', os.path.basename(fval)))

    preFile.MinMaxMetric('WPDM')  #'WPO' 'WPD'
    wpmtup = preFile.toTuple(minmax=True)
    wpmWin = preFile.Winner(minmax=True)[0]
    minmaxMargins.write_row(idx + 1, 0, wpmtup + (isCondorcet, '', os.path.basename(fval)))

    preFile.PluralityCount()
    plurtup = preFile.toTuple()
    pWin = preFile.Winner()[0]
    pluralitysheet.write_row(idx + 1, 0, plurtup + (isCondorcet, '', os.path.basename(fval)))

    preFile.BordaCount()
    bortup = preFile.toTuple()
    bWin = preFile.Winner()[0]
    bordasheet.write_row(idx + 1, 0, bortup + (isCondorcet, '', os.path.basename(fval)))

    if not (wpoWin == wpdWin == wpmWin):
        minmaxmetricComp.write_row(tmp2 + 1, 0, wpdtup, mergeformat)
        minmaxmetricComp.write(tmp2 + 1, len(wpdtup), 'Winning Votes', minmaxXoYLform)

        minmaxmetricComp.write_row(tmp2 + 2, 0, wpmtup, mergeformat)
        minmaxmetricComp.write(tmp2 + 2, len(wpmtup), 'Margins', minmaxMarginsform)

        minmaxmetricComp.write_row(tmp2 + 3, 0, wpotup, mergeformat)
        minmaxmetricComp.write(tmp2 + 3, len(wpotup), 'Pairwise Opposition', minmaxXoYform)
        minmaxmetricComp.merge_range(tmp2 + 1, len(wpotup) + 1, tmp2 + 3, len(plurtup) + 1, isCondorcet, mergeformat)
        minmaxmetricComp.merge_range(tmp2 + 1, len(wpotup) + 2, tmp2 + 3, len(plurtup) + 2, os.path.basename(fval),
                                     mergeformat)

        tmp2 += 4

    if not (cWin == bWin == pWin):
        comparesheet.write_row(tmp + 1, 0, plurtup, mergeformat)
        comparesheet.write(tmp + 1, len(plurtup), 'Plurality', plurformat)
        comparesheet.write_row(tmp + 2, 0, bortup, mergeformat)
        comparesheet.write(tmp + 2, len(bortup), 'Borda', bordformat)
        comparesheet.write_row(tmp + 3, 0, copetup, mergeformat)
        comparesheet.write(tmp + 3, len(copetup), 'Copeland', copeformat)
        comparesheet.write_row(tmp + 4, 0, wpdtup, mergeformat)
        comparesheet.write(tmp + 4, len(wpdtup), 'Winning Votes', minmaxXoYLform)

        comparesheet.write_row(tmp + 5, 0, wpmtup, mergeformat)
        comparesheet.write(tmp + 5, len(wpmtup), 'Margins', minmaxMarginsform)

        comparesheet.write_row(tmp + 6, 0, wpotup, mergeformat)
        comparesheet.write(tmp + 6, len(wpotup), 'Pairwise Opposition', minmaxXoYform)

        comparesheet.merge_range(tmp + 1, len(plurtup) + 1, tmp + 6, len(plurtup) + 1, isCondorcet, mergeformat)
        comparesheet.merge_range(tmp + 1, len(plurtup) + 2, tmp + 6, len(plurtup) + 2, os.path.basename(fval),
                                 mergeformat)
        tmp += 7

workbook.close()