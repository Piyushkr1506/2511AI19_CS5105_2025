/*Minimum Cost Climbing Stairs or Frog Jump */
class Solution {
  public:
    int minCost(vector<int>& height) {
        int n = height.size();
        vector<int> dp(n,0);
        dp[1] = abs(height[1] - height[0]);
        for(int i=2;i<n;i++){
            // ith = dp(i-2) + (height[i] - height[i-2])
            // ith = dp(i-1) + (height[i] - height[i-1])
            dp[i] = min(dp[i-2]+abs(height[i] - height[i-2]),dp[i-1] + abs(height[i] - height[i-1]));
        }
        return dp[n-1];
    }
};