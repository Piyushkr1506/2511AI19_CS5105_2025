/*Ninja's Training or Geek's Training */
class Solution {
  public:
    int maximumPoints(vector<vector<int>>& arr) {
        int size = arr.size();
        vector<vector<int>> dp(size,vector<int>(3,0));
        
        for(int i=0;i<3;i++) 
            dp[0][i] = arr[0][i];
        
        for(int i=1;i<size;i++){
            dp[i][0] = max(dp[i-1][1],dp[i-1][2]) + arr[i][0];
            dp[i][1] = max(dp[i-1][0],dp[i-1][2]) + arr[i][1];
            dp[i][2] = max(dp[i-1][0],dp[i-1][1]) + arr[i][2];
        }
        
        int max_ = INT_MIN;
        for(int i=0;i<3;i++) 
            max_ = max(max_,dp[size-1][i]);
        return max_;
    }
};