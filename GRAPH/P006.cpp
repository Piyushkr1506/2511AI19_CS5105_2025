/*
Dijkstra’s Algorithm - Using Priority Queue : G-32
Given a weighted, undirected, and connected graph of V vertices and an adjacency list adj
where adj[i] is a list of lists containing two integers where the first integer of each list j denotes
there is an edge between i and j, second integers corresponds to the weight of that edge. You
are given the source vertex S and You have to Find the shortest distance of all the vertex from
the source vertex S. You have to return a list of integers denoting the shortest distance between
each node and the Source vertex S.
Note: The Graph doesn’t contain any negative weight cycle.
Time Complexity: O(E*logV)
*/

#include<bits/stdc++.h>
using namespace std;

#define INF 99999

class Solution {
  public:
    vector<int> dijkstra(int V, vector<vector<int>> &edges, int src) {
        vector<vector<pair<int,int>>> adj(V);
        vector<int> dist(V,INF);
        for(auto edge: edges){
            int i=edge[0], j=edge[1], wt=edge[2];
            adj[i].push_back({j,wt});
            adj[j].push_back({i,wt});
        }
        priority_queue<int,vector<int>,greater<int>> pq;
        pq.push(src);
        dist[src] = 0;
        while(!pq.empty()){
            int node = pq.top(); pq.pop();
            for(auto it: adj[node]){
                int next = it.first, d = it.second;
                if(dist[node]+d<dist[next]){
                    dist[next] = dist[node] + d;
                    pq.push(next);
                }
            }
        }
        return dist;
    }
};

int main(){
    Solution sol;
    int V = 5;
    vector<vector<int>> edges = {{0, 1, 4}, {0, 2, 8}, {1, 4, 6}, {2, 3, 2}, {3, 4, 10}};
    int src = 0;
    vector<int> dist = sol.dijkstra(V, edges, src);
    for(int i=0; i<V; i++){
        cout << "Distance from source " << src << " to " << i << " is " << dist[i] << endl;
    }
    return 0;
}