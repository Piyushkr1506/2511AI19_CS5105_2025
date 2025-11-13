/*
Number of Enclaves 
Problem Statement: You are given an N x M binary matrix grid, where 0 represents a sea cell 
and 1 represents a land cell. A move consists of walking from one land cell to another adjacent 
(4-directionally) land cell or walking off the boundary of the grid. Find the number of land cells in 
the grid for which we cannot walk off the boundary of the grid in any number of moves. 
Time Complexity: O(N*M)
Space Complexity: O(N*M) for visited array
*/

#include<bits/stdc++.h>
using namespace std;

class Solution {
    vector<pair<int,int>> dirs = {{0,-1},{0,1},{-1,0},{1,0}};
  public:
    void dfs(vector<vector<int>>& grid, vector<vector<bool>> &visited,int i,int j,int n,int m){
        visited[i][j] = true;
        for(auto dir: dirs) {
            int ni = dir.first + i, nj = dir.second + j;
            if(ni>=0 && ni<n && nj>=0 && nj<m && !visited[ni][nj] && grid[ni][nj]==1){
                dfs(grid,visited,ni,nj,n,m);
                visited[ni][nj] = true;
            }
        }
    }
    int numberOfEnclaves(vector<vector<int>> &grid) {
        int count = 0;
        int n = grid.size(), m = grid[0].size();
        vector<vector<bool>> visited(n,vector<bool>(m,false));
        for(int i=0;i<n;i++){
            if(!visited[i][0] && grid[i][0]==1){
                dfs(grid,visited,i,0,n,m);
            }
            if(!visited[i][m-1] && grid[i][m-1]==1){
                dfs(grid,visited,i,m-1,n,m);
            }
        }
        for(int i=0;i<m;i++){
            if(!visited[0][i] && grid[0][i]==1){
                dfs(grid,visited,0,i,n,m);
            }
            if(!visited[n-1][i] && grid[n-1][i]==1){
                dfs(grid,visited,n-1,i,n,m);
            }
        }
        for(int i=0;i<n;i++){
            for(int j=0;j<m;j++){
                if(!visited[i][j] && grid[i][j] == 1) count += 1;
            }
        }
        return count;
    }
};

int main(){
    Solution sol;
    vector<vector<int>> grid = {
        {0, 0, 0, 0},
        {1, 0, 1, 0},
        {0, 1, 1, 0},
        {0, 0, 0, 0}
    };
    int result = sol.numberOfEnclaves(grid);
    cout << result << endl;
    return 0;
}