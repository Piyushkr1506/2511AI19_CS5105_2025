/*Partition Equal Subset Sum*/
class Solution {
  public:
    bool helper(vector<int> &arr,int idx, int size,int sum,vector<vector<int>> &dp) {
        if(sum == 0) return true;
        if(idx >= size) return false;
        if(dp[idx][sum] != -1) return dp[idx][sum];
        bool inc = false;
        if(sum >= arr[idx]) inc = helper(arr,idx+1,size,sum-arr[idx],dp);
        bool exc = helper(arr,idx+1,size,sum,dp);
        return dp[idx][sum] = (inc or exc);
    }
    bool equalPartition(vector<int>& arr) {
        int size = arr.size();
        int total_sum = accumulate(arr.begin(),arr.end(),0);
        if(total_sum%2) return false;
        total_sum /= 2;
        vector<vector<int>> dp(size+1,vector<int>(total_sum+1,-1));
        return helper(arr,0,size,total_sum,dp);
    }
};