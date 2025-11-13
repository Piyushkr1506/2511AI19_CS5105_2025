/*
Surrounded Regions | Replace O’s with X’s 
Problem Statement: Given a matrix mat of size N x M where every element is either ‘O’ or ‘X’. 
Replace all ‘O’ with ‘X’ that is surrounded by ‘X’. An ‘O’ (or a set of ‘O’) is considered to be 
surrounded by ‘X’ if there are ‘X’ at locations just below, just above just left, and just right of it.
Time Complexity: O(N*M)
Space Complexity: O(N*M) for visited array
*/

#include<bits/stdc++.h>
using namespace std;



class Solution {
    vector<pair<int,int>> dirs = {{0,-1},{0,1},{-1,0},{1,0}};
  public:
    void dfs(vector<vector<char>>& mat, vector<vector<bool>> &visited,int i,int j,int n,int m){
        visited[i][j] = true;
        for(auto dir: dirs) {
            int ni = dir.first + i, nj = dir.second + j;
            if(ni>=0 && ni<n && nj>=0 && nj<m && !visited[ni][nj] && mat[ni][nj]=='O'){
                dfs(mat,visited,ni,nj,n,m);
                visited[ni][nj] = true;
            }
        }
    }
    vector<vector<char>> fill(vector<vector<char>>& mat) {
        int n = mat.size(), m = mat[0].size();
        vector<vector<bool>> visited(n,vector<bool>(m,false));
        for(int i=0;i<n;i++){
            if(!visited[i][0] && mat[i][0]=='O'){
                dfs(mat,visited,i,0,n,m);
            }
            if(!visited[i][m-1] && mat[i][m-1]=='O'){
                dfs(mat,visited,i,m-1,n,m);
            }
        }
        for(int i=0;i<m;i++){
            if(!visited[0][i] && mat[0][i]=='O'){
                dfs(mat,visited,0,i,n,m);
            }
            if(!visited[n-1][i] && mat[n-1][i]=='O'){
                dfs(mat,visited,n-1,i,n,m);
            }
        }
        for(int i=0;i<n;i++){
            for(int j=0;j<m;j++){
                if(!visited[i][j] && mat[i][j] == 'O') mat[i][j] = 'X';
            }
        }
        return mat;
    }
};

int main(){
    Solution sol;
    vector<vector<char>> mat = {
        {'X', 'X', 'X', 'X'},
        {'X', 'O', 'O', 'X'},
        {'X', 'X', 'O', 'X'},
        {'X', 'O', 'X', 'X'}
    };
    vector<vector<char>> result = sol.fill(mat);
    for(auto row: result){
        for(auto cell: row){
            cout << cell << " ";
        }
        cout << endl;
    }
    return 0;
}