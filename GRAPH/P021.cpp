/*
Detect cycle in a directed graph (using DFS) : G 19 
Problem Statement: Given a directed graph with V vertices and E edges, check whether it 
contains any cycle or not.  
Pre-req: Graph representation, DFS
Time Complexity: O(V + E)
*/
#include<bits/stdc++.h>
using namespace std;

class Solution {
  public:
    bool dfs(int idx,vector<bool> &visited, vector<bool> &recStack,vector<vector<int>> &adj){
        if(recStack[idx]) return true;
        if(visited[idx]) return false;
        visited[idx] = true;
        recStack[idx] = true;
        for(int next: adj[idx]){
            if(dfs(next,visited,recStack,adj)) return true;
        }
        recStack[idx] = false;
        return false;
    }
    bool isCyclic(int V, vector<vector<int>> &edges) {
        vector<vector<int>> adj(V,vector<int>());
        for(auto edge: edges){
            adj[edge[0]].push_back(edge[1]);
        }
        vector<bool> visited(V,false), recStack(V,false);
        for(int i=0;i<V;i++){
            if(!visited[i] and dfs(i,visited,recStack,adj)) {
                return true;
            }
        }
        return false;
    }
};

int main(){
    Solution sol;
    vector<vector<int>> edges = {
        {0, 1}, {1, 2}, {2, 0}, {2, 3}
    };
    cout<<sol.isCyclic(4, edges)<<endl;
    return 0;
}