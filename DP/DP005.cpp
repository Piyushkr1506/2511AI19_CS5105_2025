/*House Robber II: Circular House arrangement */
class Solution {
public:
    int robLinear(vector<int> &nums,int st,int end){
        int n = end - st + 1;
        if(n == 1) return nums[st];
        vector<int> dp(n,0);
        dp[0] = nums[st];
        dp[1] = max(nums[st],nums[st+1]);
        for(int i=2;i<n;i++){
            dp[i] = max(dp[i-1], dp[i-2]+nums[st+i]);
        }
        return dp[n-1];
    }
    int rob(vector<int>& nums) {
        int size = nums.size();
        if(size == 1) return nums[0];
        int first = robLinear(nums, 0, size-2);
        int last = robLinear(nums, 1, size-1);
        return max(first,last);
    }
};