/*Path With Minimum Effort
You are a hiker preparing for an upcoming hike. You are given heights, a 2D array of size rows x
columns, where heights[row][col] represents the height of the cell (row, col). You are situated in
the top-left cell, (0, 0), and you hope to travel to the bottom-right cell, (rows-1, columns-1)
(i.e.,0-indexed). You can move up, down, left, or right, and you wish to find a route that requires
the minimum effort.
A route's effort is the maximum absolute difference in heights between two consecutive cells of
the route.
Time Complexity: O(E*logV)
*/
#include<bits/stdc++.h>
using namespace std;
#define ppii pair<int,pair<int,int>>
#define INF 99999

class Solution {
    vector<pair<int,int>> dirs = {{0,-1},{0,1},{-1,0},{1,0}};
  public:
    int MinimumEffort(int n, int m, vector<vector<int>> &heights) {
        priority_queue<ppii,vector<ppii>,greater<ppii>> pq; // {effort, {i,j}}
        pq.push({0,{0,0}});
        vector<vector<int>> dist(n,vector<int>(m,INF));
        if(n==1 && m==1) return 0;
        while(!pq.empty()){
            auto front = pq.top(); pq.pop();
            int effort = front.first, i = front.second.first, j = front.second.second;
            for(auto dir: dirs){
                int ni = i + dir.first, nj = j + dir.second;
                if(ni>=0 && ni<n && nj>=0 && nj<m && max(effort,abs(heights[ni][nj]-heights[i][j])) < dist[ni][nj]){
                    dist[ni][nj] = max(effort,abs(heights[ni][nj]-heights[i][j]));
                    pq.push({dist[ni][nj],{ni,nj}});
                }
            }
        }
        return dist[n-1][m-1];
    }
};


int main(){
    Solution sol;
    int n = 3, m = 3;
    vector<vector<int>> heights = {
        {1,2,2},{3,8,2},{5,3,5}
    };
    int result = sol.MinimumEffort(n, m, heights);
    cout << "Minimum effort: " << result << endl;
    return 0;
}