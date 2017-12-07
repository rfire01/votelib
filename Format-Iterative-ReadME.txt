Data description for iterative

Session description:
DateTime: TimeStamp of session
LabExperiment: Lab Session (Yes/No)
SessionComments: comments..

Game description:
NumberOfCandidates: Total number of candidates 
CandName1: Name of candidate 1
CandName2: Name of candidate 2
CandName3: Name of candidate 3

...
Util1: score if 1st preferred candidate is elected
Util2: score if 2nd preferred candidate is elected
Util3: score if 3rd preferred candidate is elected

GameID: Game unique identifier in session
NumVotes: Number of votes in the poll (in iterative setting this is the number of players).
TurnDeadLine: Deadline (in number of rounds) for ending the game. 
TurnLimitKnown: Round limit is common knowledge (Yes/No).
PrefsKnown:  Preferences common knowledge (Yes/No).

SessionCounter:  session counter  for a specific VoterID
GameCounterInSession: game counter for a specific voterID in session SessionCounter


Game data:
VoterID: ID of voter
Pref1: The 1st priority index of VoterID
Pref2: The 2nd priority index of VoterID
Pref3: The 3rd priority index of VoterID


..
Turn data:
TurnIDX: Current turn index in GameID
LeaderPre: Current leader(s) in poll
VotesCand1Pre: Total number of votes for candidate 1 in poll
VotesCand2Pre: Total number of votes for candidate 2 in poll
VotesCand3Pre: Total number of votes for candidate 3 in poll
Action: Vote that was casted by voterID, (value is the index in his preference list)
ActionChange: Did VoterID change his vote in TurnIDX from previous turn
ResponseTime: How long it took VoterID to cast his vote (in seconds)
LeaderPost: which candidate(s) leading after turnIDX
VotesCand1Post: number of votes candidate 1 received after turnIDX
VotesCand2Post: number of votes candidate 2 received after turnIDX
VotesCand3Post: number of votes candidate 3 received after turnIDX
..

End-Game data:
NumTurns: The total number of turns in GameID until the end of the game (at most RoundDeadline * NumVotes)
Converge: did the game converge (Yes\No)
EndWinnerCand: Which candidate(s) was elected at the end.  
EndVotesCand1: Final number of votes for candidate 1
EndVotesCand2: Final number of votes for candidate 2
EndVotesCand3: Final number of votes for candidate 3
