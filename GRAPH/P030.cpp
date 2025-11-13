/*
Dijkstra’s Algorithm - Using Set : G-33 
Given a weighted, undirected, and connected graph of V vertices and an adjacency list adj 
where adj[i] is a list of lists containing two integers where the first integer of each list j denotes 
there is an edge between i and j, second integers corresponds to the weight of that edge. You 
are given the source vertex S and You have to Find the shortest distance of all the vertex from 
the source vertex S. You have to return a list of integers denoting the shortest distance between 
each node and Source vertex S. 
Note: The Graph doesn’t contain any negative weight cycle 
Time Complexity: O(ElogV) - using set
*/
#include<bits/stdc++.h>
using namespace std;

#define INF 99999

class Solution {
  public:
    vector<int> dijkstra(int V, vector<vector<int>> &edges, int src) {
        vector<vector<pair<int,int>>> adj(V);
        vector<int> dist(V,INF);
        set<pair<int,int>> st;
        
        for(auto edge: edges){
            int i=edge[0], j=edge[1], wt=edge[2];
            adj[i].push_back({j,wt});
            adj[j].push_back({i,wt});
        }
        
        st.insert({0,src});
        dist[src] = 0;
        while(!st.empty()){
            auto begin = *(st.begin()); st.erase(begin);
            int curr_dist = begin.first, node = begin.second;
            for(auto it: adj[node]){
                int next = it.first, d = it.second;
                if(dist[node]+d<dist[next]){
                    if(dist[next] != INF) {
                        st.erase({dist[next],next});
                    }
                    dist[next] = dist[node] + d;
                    st.insert({dist[next],next});
                }
            }
        }
        return dist;
    }
};

int main(){
    Solution sol;
    int n = 5, m = 6 ;
    vector<vector<int>> edges ={{0,1,2},{0,2,4},{1,2,1},{1,3,7},{2,4,3},{3,4,1}} ;
    vector<int> dist = sol.dijkstra(n,edges,0);
    for(int d: dist) cout<<d<<" ";
    return 0;
}