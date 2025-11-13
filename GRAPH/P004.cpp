/*Distance of Nearest Cell having 1
Problem Statement: Given a binary grid of N*M. Find the distance of the nearest 1 in the grid for
each cell.
The distance is calculated as |i1 - i2| + |j1 - j2|, where i1, j1 are the row number and column
number of the current cell, and i2, j2 are the row number and column number of the nearest cell
having value 1
*/
#include<bits/stdc++.h>
using namespace std;

class Solution {
    vector<pair<int,int>> dirs = {{0,-1},{0,1},{-1,0},{1,0}};
  public:
    vector<vector<int>> nearest(vector<vector<int>>& grid) {
        int n = grid.size(), m = grid[0].size(); 
        queue<pair<int,int>> q;
        for(int i=0;i<n;i++){
            for(int j=0;j<m;j++){
                if(grid[i][j] == 1) {
                    q.push({i,j});
                    grid[i][j] = -1;
                }
            }
        }
        int dist = 1;
        while(!q.empty()){
            int size = q.size();
            for(int i=0;i<size;i++){
                auto curr = q.front(); q.pop();
                int ci = curr.first, cj = curr.second;
                for(auto dir: dirs){
                    int ni = dir.first + ci, nj = dir.second + cj;
                    if(ni>=0 && ni<n && nj>=0 && nj<m && grid[ni][nj]==0){
                        grid[ni][nj] = dist;
                        q.push({ni,nj});
                    }
                }
            }
            dist += 1;
        }
        for(int i=0;i<n;i++){
            for(int j=0;j<m;j++){
                if(grid[i][j] == -1) grid[i][j] = 0;
            }
        }
        return grid;
    }
};

int main(){
    Solution sol;
    vector<vector<int>> grid = {
        {1, 0, 1},
        {1, 1, 0},
        {1, 0, 0}
    };
    vector<vector<int>> result = sol.nearest(grid);
    for(auto row: result){
        for(auto cell: row){
            cout << cell << " ";
        }
        cout << endl;
    }
    return 0;
}