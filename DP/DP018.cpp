/*Target Sum*/
class Solution {
    const int mod = 1e9+7;  
  public:
    long long countPartitions(vector<int>& arr, int d) {
        int t = accumulate(arr.begin(),arr.end(),0);
        if((t+d)%2 != 0 or t<abs(d)) return 0;
        int target = (t+d)/2;
        int size = arr.size();
        vector<vector<long long>> dp(size+1,vector<long long>(target+1,0));
        for(int i=0;i<=size;i++) 
            dp[i][0] = 1;
        for(int i=1;i<=size;i++){
            for(int sum=0;sum<=target;sum++){
                if(sum>=arr[i-1]) 
                    dp[i][sum] = (dp[i-1][sum-arr[i-1]] + dp[i-1][sum])%mod;
                else 
                    dp[i][sum] = dp[i-1][sum];
            }
        }
        return dp[size][target];
    }
    long long findTargetSumWays(int n, vector<int>& A, int target) {
        long long res = countPartitions(A,target);
        return res;
    }
};