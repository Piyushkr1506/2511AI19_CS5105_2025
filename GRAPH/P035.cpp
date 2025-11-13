/*
Find the City With the Smallest Number of Neighbours at a Threshold Distance 
Problem Statement: There are n cities numbered from 0 to n-1. Given the array edges where 
edges[i] = [fromi, toi,weighti]  represents a bidirectional and weighted edge between cities fromi 
and toi, and given the integer distance Threshold. You need to find out a city with the smallest 
number of cities that are reachable through some path and whose distance is at most Threshold 
Distance, If there are multiple such cities, our answer will be the city with the greatest number. 
Note: that the distance of a path, connecting cities i and j are equal to the sum of the edges' 
weights along that path.
Time Complexity: O(n^3) - Using Floyd-Warshall Algorithm, O(n^2*logn) - Using Dijkstra's Algorithm
Space Complexity: O(n^2) - For the distance matrix

*/

#include<bits/stdc++.h>
using namespace std;

class Solution {
  public:
    /* Using Floyd-Warshall Algorithm
    int findCity(int n, int m, vector<vector<int>>& edges, int threshold) {
        vector<vector<int>> dp(n,vector<int>(n,1e9));
        for(const auto edge: edges){
            auto st = edge[0], end = edge[1] , wt = edge[2];
            dp[st][end] = wt;
            dp[end][st] = wt;
        }
        for(int i=0;i<n;i++) dp[i][i] = 0;
        for(int mid=0;mid<n;mid++){
            for(int st=0;st<n;st++){
                for(int end=0;end<n;end++){
                    if(dp[st][end] > (dp[st][mid] + dp[mid][end])) {
                        dp[st][end] = dp[st][mid] + dp[mid][end];
                    }
                }
            }
        }
        int idx = -1, min_count = 1e9;
        for(int i=0;i<n;i++){
            int count = 0;
            for(int dist: dp[i]){
                if(dist <= threshold) {
                    count += 1;
                }
            }
            if(min_count >= count) {
                idx = i;
                min_count = count;
            }
        }
        return idx;
    }
    */
    vector<int> dijkstra(vector<vector<int>> &graph, int src){
        int size = graph.size();
        vector<int> dist(size,1e9);
        priority_queue<int,vector<int>,greater<int>> pq;
        
        pq.push(src);
        dist[src] = 0;
        
        while(!pq.empty()){
            int front = pq.top(); pq.pop();
            for(int i=0;i<size;i++){
                if(graph[front][i] != 1e9 and dist[i] > dist[front]+graph[front][i]) {
                    dist[i] = dist[front]+graph[front][i];
                    pq.push(i);
                }
            }
        }
        
        return dist;
    }
    int findCity(int n, int m, vector<vector<int>>& edges, int threshold) {
        int idx = -1, min_count = 1e9;
        vector<vector<int>> graph(n,vector<int>(n,1e9));
        
        for(auto edge: edges) {
            auto st = edge[0], end = edge[1] , wt = edge[2];
            graph[st][end] = wt;
            graph[end][st] = wt;
        }
        
        for(int i=0;i<n;i++){
            vector<int> distance = dijkstra(graph,i);
            int count = 0;
            for(int dist: distance){
                if(dist <= threshold) {
                    count += 1;
                }
            }
            if(min_count >= count) {
                idx = i;
                min_count = count;
            }
        }
        return idx;
    }
};

int main(){
    Solution sol;
    int n = 4, m = 4;
    vector<vector<int>> edges = {{0,1,3},{1,2,1},{1,3,4},{2,3,1}};
    int threshold = 4;
    cout<<sol.findCity(n,m,edges,threshold)<<endl;
    return 0;
}