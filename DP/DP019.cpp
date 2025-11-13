/*Wildcard String matching*/
class Solution {
  public:
    bool helper(string wild, string pattern, int i, int j,vector<vector<int>> &dp) {
        if(i<0 and j<0) 
            return true;
        if(i<0 and j>=0) 
            return false;
        if(i>=0 and j<0) {
             for (int k = 0; k <= i; k++) {
                if (wild[k] != '*') 
                    return false;
            }
            return true;
        }
        if(dp[i][j] != -1) 
            return dp[i][j];
        if((wild[i] == pattern[j]) or wild[i]=='?') 
            return dp[i][j] = helper(wild,pattern,i-1,j-1,dp);
        if(wild[i] == '*') 
            return dp[i][j] = helper(wild,pattern,i-1,j,dp) or helper(wild,pattern,i,j-1,dp);
        return false;
    }
    bool match(string wild, string pattern) {
        int m = wild.length(), n = pattern.length();
        vector<vector<int>> dp(m+1,vector<int>(n+1,-1));
        return helper(wild,pattern,m-1,n-1,dp);
    }
};