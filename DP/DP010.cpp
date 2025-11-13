/*Triangle Path Sum */
class Solution {
  public:
    int minPathSum(vector<vector<int>>& triangle) {
        int m = triangle.size();
        vector<vector<int>> dp(m);
        for(int i=0;i<m;i++){
            dp[i] = vector<int>(i+1);
        }
        dp[0][0] = triangle[0][0];
        for(int i=1;i<m;i++){
            for(int j=0;j<=i;j++){
                int min_ = INT_MAX;
                if(j<i) min_ = min(min_,dp[i-1][j]);
                if(j>=1) min_ = min(min_,dp[i-1][j-1]);
                dp[i][j] = min_ + triangle[i][j];
            }
        }
        return *min_element(dp[m-1].begin(),dp[m-1].end());
    }
};