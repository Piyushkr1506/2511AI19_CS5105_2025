/*
Shortest Distance in a Binary Maze
Problem Statement:
Given an n * m matrix grid where each element can either be 0 or 1. You need to find the
shortest distance between a given source cell to a destination cell. The path can only be created
out of a cell if its value is 1.
If the path is not possible between the source cell and the destination cell, then return -1.
Note: You can move into an adjacent cell if that adjacent cell is filled with element 1. Two cells
are adjacent if they share a side. In other words, you can move in one of four directions, Up,
Down, Left, and Right
Time Complexity: BFS: O(N * M), DFS: O(4^(N*M))
*/

#include<bits/stdc++.h>
using namespace std;


class Solution {
    vector<pair<int,int>> dirs = {{0,-1},{0,1},{-1,0},{1,0}};
  public:
    int shortestDistance(int n, int m, vector<vector<int>> mat, int X, int Y) {
        if(mat[0][0]==0 or mat[X][Y]==0) return -1;
        vector<vector<int>> visited(n, vector<int>(m,false));
        int dist = 0;
        queue<pair<int,int>> q;
        q.push({0,0});
        while(!q.empty()){
            int size = q.size();
            for(int i=0;i<size;i++){
                auto front = q.front(); q.pop();
                int ci = front.first, cj = front.second;
                if(ci==X && cj==Y) return dist;
                visited[ci][cj] = true;
                for(auto dir: dirs){
                    int ni = ci + dir.first, nj = cj + dir.second;
                    if(ni>=0 && ni<n && nj>=0 && nj<m && mat[ni][nj]==1 && !visited[ni][nj]){
                        q.push({ni,nj});
                        visited[ni][nj] = true;
                    }
                }
            }
            dist += 1;
        }
        return -1;
    }
};

int main(){
    Solution sol;
    int n = 3, m = 4;
    vector<vector<int>> mat = {
        {1,0,0,0}, 
        {1,1,0,1},
        {0,1,1,1}
    };
    int X = 2, Y = 3;
    int result = sol.shortestDistance(n, m, mat, X, Y);
    cout << "Shortest distance: " << result << endl;
    return 0;
}