// Partitions with Given Difference
class Solution {
  public:
    int countPartitions(vector<int>& arr, int d) {
        int t = accumulate(arr.begin(),arr.end(),0);
        if((t+d)%2 != 0 or t<d) return 0;
        int target = (t+d)/2;
        int size = arr.size();
        vector<vector<int>> dp(size+1,vector<int>(target+1,0));
        for(int i=0;i<=size;i++) dp[i][0] = 1;
        for(int i=1;i<=size;i++){
            for(int sum=0;sum<=target;sum++){
                if(sum>=arr[i-1]) 
                    dp[i][sum] = dp[i-1][sum-arr[i-1]] + dp[i-1][sum];
                else 
                    dp[i][sum] = dp[i-1][sum];
            }
        }
        return dp[size][target];
    }
};