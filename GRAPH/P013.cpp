/*
Flood Fill Algorithm - Graphs
Problem Statement: An image is represented by a 2-D array of integers, each integer
representing the pixel value of the image. Given a coordinate (sr, sc) representing the starting
pixel (row and column) of the flood fill, and a pixel value newColor, "flood fill" the image.
To perform a "flood fill", consider the starting pixel, plus any pixels connected 4-directionally to
the starting pixel of the same colour as the starting pixel, plus any pixels connected
4-directionally to those pixels (also with the same colour as the starting pixel), and so on.
Replace the colour of all of the aforementioned pixels with the newColor.
Pre-req: Connected Components, Graph traversal techniques
Time Complexity: O(M*N) for both BFS and DFS
Space Complexity: O(M*N) for both BFS and DFS
*/

#include<bits/stdc++.h>
using namespace std;

class Solution {
    vector<pair<int,int>> dirs = {{0,1},{1,0},{0,-1},{-1,0}};
  public:
    void dfs(int sr, int sc,vector<vector<int>>& image,vector<vector<bool>> &visited, int newColor, int oldColor,int n,int m){
        visited[sr][sc] = true;
        image[sr][sc] = newColor;
        for(auto dir: dirs){
            int nr = sr + dir.first, nc = sc + dir.second;
            if(nr>=0 and nr<n and nc>=0 and nc<m and image[nr][nc]==oldColor and !visited[nr][nc]){
                dfs(nr,nc,image,visited,newColor,oldColor,n,m);
                visited[nr][nc] = true;
            }
        }
    }
    
    vector<vector<int>> floodFill(vector<vector<int>>& image, int sr, int sc, int newColor) {
        int n = image.size(), m = image[0].size();
        vector<vector<bool>> visited(n,vector<bool>(m,false));
        dfs(sr,sc,image,visited,newColor,image[sr][sc],n,m);
        return image;
    }
};

int main(){
    Solution sol;
    vector<vector<int>> image = {
        {1, 1, 1, 0}, {0, 1, 1, 1}, {1, 0, 1, 1}
    };
    int sr = 1, sc = 2, newColor = 2;
    vector<vector<int>> res = sol.floodFill(image,sr,sc,newColor);
    for(auto row: res){
        for(auto it: row) cout<<it<<" ";
        cout<<endl;
    }
    return 0;
}