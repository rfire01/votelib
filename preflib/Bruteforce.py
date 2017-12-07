import math
import itertools
import string
import xlsxwriter
from majoritycount import prefLib

__author__ = 'maortal'
NVOTERS = 10


class prefBruthForce:
    def __init__(self, nCands, nVoters):
        self.nCands = nCands
        self.nVoters = nVoters
        self.emptyPrefList = []
        for perm in itertools.permutations(xrange(1, 4)):
            l = [str(x) for x in perm]
            self.emptyPrefList.append(','.join(l))

    def VotingDistribution(self):
        return self._votingDistribution(self.nVoters, [self.nVoters] * math.factorial(self.nCands))

    def _votingDistribution(self, voters, prefPermutation):
        if not voters:
            yield len(prefPermutation) * (0,)
        elif len(prefPermutation) == 1:
            yield (voters,)
        else:
            for voters_for_first_pref in xrange(voters, -1, -1):
                votes_in_other_prefs = voters - voters_for_first_pref
                for distribution_other in self._votingDistribution(votes_in_other_prefs, prefPermutation[1:]):
                    yield (voters_for_first_pref,) + distribution_other

    def prefDest(self):
        res = []
        res.append(str(self.nCands))
        for idx, ch in enumerate(string.uppercase[:self.nCands]):
            res.append(str(idx + 1) + ',' + ch)
        res.append(str(self.nVoters) + ',' + str(self.nVoters) + ',' + str(math.factorial(self.nCands)))
        prefpoint = len(res)
        for perf in self.VotingDistribution():
            res = res[:prefpoint]
            for idx, x in enumerate(perf):
                res.append(str(x) + "," + self.emptyPrefList[idx])
            yield perf, res


if __name__ == "__main__":


    """
    Excel File design
    """
    workbook = xlsxwriter.Workbook('BF-Results.xlsx')
    plurformat = workbook.add_format({'font_color': 'red'})
    bordformat = workbook.add_format({'font_color': 'green'})
    copeformat = workbook.add_format({'font_color': 'blue'})
    minmaxXoYform = workbook.add_format({'font_color': '#B45F04'})
    minmaxXoYLform = workbook.add_format({'font_color': 'orange'})
    minmaxMarginsform = workbook.add_format({'font_color': '#0489B1'})
    mergeformat = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

    possibImportant = workbook.add_format({'bg_color': 'yellow', 'bold': True})

    candsRow3 = ('Candidate 1', 'Candidate 2', 'Candidate 3', 'Winner', 'Total Votes', 'Method', 'Has Condorcet?',
                 '','','Configuration')

    for numVoters in xrange(3, NVOTERS):
        x = prefBruthForce(3, numVoters)
        currentWS = workbook.add_worksheet(str(numVoters) + ' Voters')
        currentWS.write_row('A1', candsRow3)
        tmp=0
        for perf, y in x.prefDest():
            preFile = prefLib(y, perf)

            preFile.Copeland()
            copetup = preFile.toTuple()
            cWin, chkCondorcet = preFile.Winner()
            isCondorcet = True if chkCondorcet == (preFile.numCands - 1) else False

            preFile.Minimax()
            preFile.MinMaxMetric('WPO')  #'WPD' 'WPDM'
            wpotup = preFile.toTuple(minmax=True)
            wpoWin = preFile.Winner(minmax=True)[0]

            preFile.MinMaxMetric('WPD')  #'WPO' 'WPDM'
            wpdtup = preFile.toTuple(minmax=True)
            wpdWin = preFile.Winner(minmax=True)[0]

            preFile.MinMaxMetric('WPDM')  #'WPO' 'WPD'
            wpmtup = preFile.toTuple(minmax=True)
            wpmWin = preFile.Winner(minmax=True)[0]

            preFile.PluralityCount()
            plurtup = preFile.toTuple()
            pWin = preFile.Winner()[0]

            preFile.BordaCount()
            bortup = preFile.toTuple()
            bWin = preFile.Winner()[0]


            winners = [cWin, bWin, pWin, wpoWin, wpdWin, wpmWin]
            setwinners = set(winners)
            if not (cWin == bWin == pWin == wpoWin == wpdWin == wpmWin):
                currentWS.write_row(tmp + 1, 0, plurtup, mergeformat)
                currentWS.write(tmp + 1, len(plurtup), 'Plurality', plurformat)
                currentWS.write_row(tmp + 2, 0, bortup, mergeformat)
                currentWS.write(tmp + 2, len(bortup), 'Borda', bordformat)
                currentWS.write_row(tmp + 3, 0, copetup, mergeformat)
                currentWS.write(tmp + 3, len(copetup), 'Copeland', copeformat)
                currentWS.write_row(tmp + 4, 0, wpdtup, mergeformat)
                currentWS.write(tmp + 4, len(wpdtup), 'Winning Votes', minmaxXoYLform)

                currentWS.write_row(tmp + 5, 0, wpmtup, mergeformat)
                currentWS.write(tmp + 5, len(wpmtup), 'Margins', minmaxMarginsform)

                currentWS.write_row(tmp + 6, 0, wpotup, mergeformat)
                currentWS.write(tmp + 6, len(wpotup), 'Pairwise Opposition', minmaxXoYform)

                currentWS.merge_range(tmp + 1, len(plurtup) + 1, tmp + 6, len(plurtup) + 1, isCondorcet, mergeformat)

                for idx,currentconf in enumerate(preFile.PrefList):
                    if len(setwinners) > 2:
                        currentWS.write('J' + str(tmp+idx+2),currentconf[0] + ' - ' + str(currentconf[1:]),possibImportant)
                    else:
                        currentWS.write('J' + str(tmp+idx+2),currentconf[0] + ' - ' + str(currentconf[1:]))

                tmp += 7

workbook.close()
