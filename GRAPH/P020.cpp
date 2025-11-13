/*
Number of Islands 
Problem Statement: Given a grid of size NxM (N is the number of rows and M is the number of 
columns in the grid) consisting of '0's (Water) and ‘1's(Land). Find the number of islands. 
Note: An island is surrounded by water and is formed by connecting adjacent lands horizontally 
or vertically or diagonally i.e., in all 8 directions. 
Pre-req: Connected Components, Graph traversal techniques
Time Complexity: O(N*M)
Space Complexity: O(N*M) for visited array, Space can be reduced by changing the input grid from '1' to '0'
*/

#include<bits/stdc++.h>
using namespace std;

/* SPACE OPTIMIZED
class Solution {
    vector<pair<int,int>> dirs = {{0,-1},{0,1},{-1,0},{1,0}};
  public:
    void dfs(vector<vector<char>>& grid, int i,int j,int n,int m){
        grid[i][j] = '0';
        for(auto dir: dirs) {
            int ni = dir.first + i, nj = dir.second + j;
            if(ni>=0 && ni<n && nj>=0 && nj<m && grid[ni][nj]=='1'){
                dfs(grid,ni,nj,n,m);
            }
        }
    }
    int numIslands(vector<vector<char>>& mat) {
        int n = mat.size(), m = mat[0].size();
        int count = 0;
        for(int i=0;i<n;i++){
            for(int j=0;j<m;j++){
                if(mat[i][j]=='1'){
                    count += 1;
                    dfs(mat,i,j,n,m);
                }
            }
        }
        return count;
    }
};*/

class Solution {
    vector<pair<int,int>> dirs = {{-1, -1}, {-1, 0}, {-1, 1}, {1, -1},
        {1, 0}, {1, 1}, {0, 1}, {0, -1}};
  public:
    void dfs(vector<vector<char>>& grid, vector<vector<bool>> &visited,int i,int j,int n,int m){
        visited[i][j] = true;
        grid[i][j] = 'W';
        for(auto dir: dirs) {
            int ni = dir.first + i, nj = dir.second + j;
            if(ni>=0 && ni<n && nj>=0 && nj<m && !visited[ni][nj] && grid[ni][nj]=='L'){
                dfs(grid,visited,ni,nj,n,m);
                visited[ni][nj] = true;
            }
        }
    }
    int countIslands(vector<vector<char>>& mat) {
        int n = mat.size(), m = mat[0].size();
        int count = 0;
        vector<vector<bool>> visited(n,vector<bool>(m,false));
        for(int i=0;i<n;i++){
            for(int j=0;j<m;j++){
                if(!visited[i][j] && mat[i][j]=='L'){
                    dfs(mat,visited,i,j,n,m);
                    mat[i][j] = 'L';
                }
            }
        }
        for(int i=0;i<n;i++){
            for(int j=0;j<m;j++){
                if(mat[i][j] == 'L') count += 1;
            }
        }
        return count;
    }
};

int main(){
    Solution sol;
    vector<vector<char>> grid = {
        {'W', 'L', 'W', 'W', 'L'},
        {'L', 'L', 'W', 'L', 'L'},
        {'W', 'L', 'W', 'W', 'W'},
        {'W', 'L', 'L', 'W', 'W'}
    };
    int result = sol.countIslands(grid);
    cout << result << endl;
    return 0;
}