/*
Floyd Warshall Algorithm: G-42 
Problem Statement: The problem is to find the shortest distances between every pair of vertices 
in a given edge-weighted directed graph. The graph is represented as an adjacency matrix of 
size n*n. Matrix[i][j] denotes the weight of the edge from i to j. If Matrix[i][j]=-1, it means there is 
no edge from i to j. 
Do it in place. 
Time Complexity: O(n^3) where n is the number of vertices in the graph.
Space Complexity: O(1) as no extra space is used.
*/

#include<bits/stdc++.h>
using namespace std;

class Solution {
  public:
    void shortestDistance(vector<vector<int>>& mat) {
        int n = mat.size();
        for(int i=0;i<n;i++){
            for(int j=0;j<n;j++) {
                if(mat[i][j]==-1) mat[i][j] = 1e9;
                if(i==j) mat[i][j] = 0;
            }
        }
        for(int mid=0;mid<n;mid++){
            for(int st=0;st<n;st++){
                for(int end=0;end<n;end++){
                    mat[st][end] = min(mat[st][end],mat[st][mid]+mat[mid][end]);
                }
            }
        }
        for(int i=0;i<n;i++){
            for(int j=0;j<n;j++) if(mat[i][j]==1e9) mat[i][j] = -1;
        }
    }
};
int main(){
    Solution sol;
    vector<vector<int>> mat = {
        {0, 3, -1, 7},
        {8, 0, 2, -1},
        {5, -1, 0, 1},
        {2, -1, -1, 0}
    };
    sol.shortestDistance(mat);
    cout << "Shortest distance matrix:" << endl;
    for(const auto& row : mat) {
        for(const auto& val : row) {
            cout << val << " ";
        }
        cout << endl;
    }
    return 0;
}