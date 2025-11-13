/*
Detect Cycle in an Undirected Graph (using BFS)
Problem Statement: Given an undirected graph with V vertices and E edges, check whether it
contains any cycle or not.
Time Complexity: O(V + E) where V is the number of vertices and E is the number of edges
Space Complexity: O(V) for the visited array and the queue
*/

#include<bits/stdc++.h>
using namespace std;

class Solution {
  public:
    bool dfs(vector<vector<int>> &adj,vector<bool> &visited,int node,int parent){
        visited[node] = true;
        for(int next: adj[node]){
            if(!visited[next]){
                if(dfs(adj, visited, next, node)) return true;
            }else{
                if(next != parent) return true;
            }
        }
        return false;
    }
        
    
    bool isCycle(int v, vector<vector<int>>& edges) {
        vector<bool> visited(v,false);
        vector<vector<int>> adj(v);
        
        for(auto edge: edges){
            int u = edge[0], v = edge[1];
            adj[u].push_back(v);
            adj[v].push_back(u);
        }
        
        for(int i=0;i<v;i++){
            if(!visited[i]){
                if(dfs(adj,visited,i,-1)) return true;
            }
        }
        return false;
    }
};

int main(){
    Solution sol;
    int v = 4;
    vector<vector<int>> edges = {
        {0, 1}, {0, 2}, {1, 2}, {2, 3}
    };
    bool res = sol.isCycle(v,edges);
    if(res) cout<<"Cycle Present";
    else cout<<"No Cycle Present";
    return 0;
}