/*
Bellman Ford Algorithm 
Problem Statement: Given a weighted, directed and connected graph of V vertices and E 
edges, Find the shortest distance of all the vertices from the source vertex S. 
Note: If the Graph contains a negative cycle then return an array consisting of only -1.
Time Complexity: O(V * E) where V is the number of vertices and E is the number of edges.
Space Complexity: O(V) for the distance array.
*/
#include<bits/stdc++.h>
using namespace std;

// User function Template for C++
#include<climits>

class Solution {
  public:
    vector<int> bellmanFord(int V, vector<vector<int>>& edges, int src) {
        vector<int> dist(V,1e8);
        dist[src] = 0;
        for(int i=0;i<V-1;i++){
            for(const auto edge: edges){
                int u = edge[0], v = edge[1], wt = edge[2];
                if(dist[u]!=1e8 && dist[u]+wt < dist[v]){
                    dist[v] = dist[u] + wt;
                }
            }
        }
        /*Another relaxation to check for negative edges*/
        for(const auto edge: edges){
            int u = edge[0], v = edge[1], wt = edge[2];
            if(dist[u]!=1e8 && dist[u]+wt < dist[v])
                return {-1};
        }
        return dist;
    }
};

int main(){
    Solution sol;
    int V = 5, E = 8;
    vector<vector<int>> edges = {{0,1,-1},{0,2,4},{1,2,3},{1,3,2},{1,4,2},{3,2,5},{3,1,1},{4,3,-3}};
    int src = 0;
    vector<int> result = sol.bellmanFord(V, edges, src);
    if(result.size() == 1 && result[0] == -1) {
        cout << "Graph contains a negative weight cycle" << endl;
    } else {
        cout << "Vertex Distance from Source (" << src << "):" << endl;
        for(int i = 0; i < result.size(); i++) {
            cout << "Vertex " << i << ": " << result[i] << endl;
        }
    }
    return 0;
}