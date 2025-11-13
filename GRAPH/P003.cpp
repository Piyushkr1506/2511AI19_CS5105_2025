/**
Rotten Oranges : Min time to rot all oranges : BFS
Problem Statement: You will be given an m x n grid, where each cell has the following values :
2 - represents a rotten orange
1 - represents a Fresh orange
0 - represents an Empty Cell
Every minute, if a Fresh Orange is adjacent to a Rotten Orange in 4-direction ( upward,
downwards, right, and left ) it becomes Rotten.
Return the minimum number of minutes required such that none of the cells has a Fresh
Orange. If it's not possible, return -1.
Time Complexity: O(N * M)
*/

#include<bits/stdc++.h>
using namespace std;

class Solution {
    vector<pair<int,int>> dirs = {{0,-1},{0,1},{-1,0},{1,0}};
public:
    int orangesRotting(vector<vector<int>>& grid) {
        int timer = 0;
        int m = grid.size();
        int n = grid[0].size();
        vector<vector<bool>> visited(m,vector<bool>(n,false));
        queue<pair<int,int>> q;
        bool fresh = false;
        for(int i=0;i<m;i++){
            for(int j=0;j<n;j++){
                if(grid[i][j]==2){
                    q.push({i,j});
                }
                if(grid[i][j]==1){
                    fresh = true;
                }
            }
        }
        if(!fresh) {
            return 0;
        }
        while(!q.empty()){
            int qsize = q.size();
            for(int k=0;k<qsize;k++){
                auto curr = q.front(); q.pop();
                int ci = curr.first, cj = curr.second;
                visited[ci][cj] = true;
                for(auto dir: dirs){
                    int ni = ci + dir.first, nj = cj + dir.second;
                    if(ni<m && nj<n && ni>=0 && nj>=0 && visited[ci][cj] && grid[ni][nj] == 1){
                        grid[ni][nj] = 2;
                        q.push({ni,nj});
                        visited[ni][nj] = true;
                    }
                }
            }
            timer += 1;
        }
        for(int i=0;i<m;i++){
            for(int j=0;j<n;j++){
                if(grid[i][j]==1) return -1;
            }
        }
        return timer-1;
    }
};

int main(){
    Solution sol;
    vector<vector<int>> grid = {
        {2, 1, 1},
        {1, 1, 0},
        {0, 1, 1}
    };
    cout << sol.orangesRotting(grid) << endl;
    return 0;
}