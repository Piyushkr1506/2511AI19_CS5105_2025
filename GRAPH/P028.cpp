/*
Shortest Path in Undirected Graph with unit distance 
Given an Undirected Graph having unit weight, find the shortest path from the source to all other 
nodes in this graph. In this problem statement, we have assumed the source vertex to be ‘0’. If a 
vertex is unreachable from the source node, then return -1 for that vertex.
Time Complexity: O((V+E)logV) - dijkstra's algorithm, but here we can do it in O(V+E) using BFS as all edges have same weight.
*/

#include<bits/stdc++.h>
using namespace std;

class Solution {
  public:
    const int INF = 1e9+5;
    vector<int> shortestPath(vector<vector<int>>& adj, int src) {
        int V = adj.size();
        vector<int> dist(V,-1);
        dist[src] = 0;
        // priority_queue<pair<int,int>,vector<pair<int,int>>,greater<pair<int,int>>> pq;
        // pq.push({0,src});
        
        // while(!pq.empty()){
        //     auto [curr_dist,curr] = pq.top(); pq.pop();
        //     for(int next: adj[curr]){
        //         if(curr_dist + 1 < dist[next]) {
        //             dist[next] = curr_dist + 1;
        //             pq.push({dist[next],next});
        //         }   
        //     }
        // }
        // for(int i=0;i<V;i++) 
        //     if(dist[i] == INF) 
        //         dist[i] = -1;
        
        queue<int> q;
        vector<bool> visited(V,false);
        q.push(src);
        visited[src] = true;
        
        while(!q.empty()){
            int size = q.size();
            for(int i=0;i<size;i++){
                int curr = q.front(); q.pop();
                int curr_dist = dist[curr];
                for(int next: adj[curr]){
                    if(!visited[next]){
                        dist[next] = curr_dist + 1;
                        visited[next] = true;
                        q.push(next);
                    }
                }
            }
        }
        return dist;
    }
};

int main(){
    Solution sol;
    int n = 8, m = 10 ;
    vector<vector<int>> edges ={{1,0},{2,1},{0,3},{3,7},{3,4},{7,4},{7,6},{4,5},{4,6},{6,5}} ;
    int src=0 ;
    vector<vector<int>> adj(n);
    for(auto &e: edges){
        adj[e[0]].push_back(e[1]);
        adj[e[1]].push_back(e[0]);
    }
    vector<int> ans = sol.shortestPath(adj,src);
    for(int i: ans) cout<<i<<" ";
    return 0;
}