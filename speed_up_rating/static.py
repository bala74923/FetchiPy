contest_info_query= '''
                           query userContestRankingInfo($username: String!) {
  userContestRanking(username: $username) {
    attendedContestsCount
    rating
    globalRanking
    totalParticipants
    topPercentage
    badge {
      name
    }
  }
  userContestRankingHistory(username: $username) {
    attended
    trendDirection
    problemsSolved
    totalProblems
    finishTimeInSeconds
    rating
    ranking
    contest {
      title
      startTime
    }
  }
}
                           '''

solved_count_query= '''
                query getUserProfile($username: String!) {
                    allQuestionsCount {
                    difficulty
                    count
                    }
                    matchedUser(username: $username) {
                    username
                    submitStats {
                        acSubmissionNum {
                        difficulty
                        count
                        }
                    }
                    }
                }
                '''