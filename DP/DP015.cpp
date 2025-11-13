/*Minimum sum partition*/
class Solution {
  public:
    int helper(vector<int>& arr,int idx, int sum, int total_sum,vector<vector<int>> &dp){
        if(idx == 0) {
            return abs(total_sum-sum-sum);
        }
        if(dp[idx][sum] != -1) return dp[idx][sum];
        int inc = helper(arr,idx-1,sum+arr[idx-1],total_sum,dp);
        int exc = helper(arr,idx-1,sum,total_sum,dp);
        return dp[idx][sum] = min(inc,exc);
    }
    int minDifference(vector<int>& arr) {
        int size = arr.size();
        int total_sum = accumulate(arr.begin(),arr.end(),0);
        vector<vector<int>> dp(size+1,vector<int>(total_sum+1,-1));
        return helper(arr,size,0,total_sum,dp);
    }
};
