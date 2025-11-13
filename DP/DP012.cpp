/*Chocolate/Cherry Pickup*/
class Solution {
  public:
    int helper(vector<vector<int>>& grid,int cr,int cc1, int cc2, vector<vector<vector<int>>> &dp, int n, int m){
        if(cc1<0 or cc1>=m or cc2<0 or cc2>=m) {
            return -1e8;
        }
        if(cr == n-1) {
            if(cc1 == cc2) {
                return grid[cr][cc1];
            }else{
                return grid[cr][cc1] + grid[cr][cc2];
            }
        }
        if(dp[cr][cc1][cc2] != -1) {
            return dp[cr][cc1][cc2];
        }
        int max_chocolate = -1e8;
        for(int i=-1;i<=1;i++){
            for(int j=-1;j<=1;j++){
                int choco = 0;
                if(cc1 == cc2) {
                    choco = grid[cr][cc1];
                }else{
                    choco = grid[cr][cc1] + grid[cr][cc2];
                }
                choco += helper(grid,cr+1,cc1+i,cc2+j,dp,n,m);
                max_chocolate = max(max_chocolate,choco);
            }
        }
        return dp[cr][cc1][cc2] = max_chocolate;
    }
    int solve(int n, int m, vector<vector<int>>& grid) {
        vector<vector<vector<int>>> dp(n,vector<vector<int>>(m,vector<int>(m,-1)));
        int cr = 0, cc1 = 0, cc2 = m-1;
        return helper(grid,cr,cc1,cc2,dp,n,m);
    }
};