/*House Robber or Maximum adjacent sum */
class Solution {
  public:
    int findMaxSum(vector<int>& arr) {
        int n = arr.size();
        int dp[n+1] = {0};
        dp[1] = arr[0];
        for(int i=2;i<=n;i++){
            dp[i] = max(dp[i-2]+arr[i-1],dp[i-1]);
        }
        return dp[n];
    }
};