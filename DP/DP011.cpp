/*Maximum path sum in matrix*/
class Solution {
  public:
    int maximumPath(vector<vector<int>>& mat) {
        int m = mat.size();
        int n = mat[0].size();
        vector<vector<int>> dp(m,vector<int>(n));
        for(int i=0;i<n;i++){
            dp[0][i] = mat[0][i];
        }
        for(int i=1;i<m;i++){
            for(int j=0;j<n;j++){
                int max_ = dp[i-1][j];
                if(j>=1) max_ = max(max_,dp[i-1][j-1]);
                if(j<n-1) max_ = max(max_, dp[i-1][j+1]);
                dp[i][j] = max_ + mat[i][j];
            }
        }
        return *max_element(dp[m-1].begin(),dp[m-1].end());
    }
};