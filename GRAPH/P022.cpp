/*
Topological Sort Algorithm 
Problem Statement: Given a Directed Acyclic Graph (DAG) with V vertices and E edges, Find 
any Topological Sorting of that Graph. 
Note: In topological sorting, node u will always appear before node v if there is a directed edge 
from node u towards node v(u -> v).
Time Complexity: O(V + E)
Space Complexity: O(V) for stack and visited array
*/

#include<bits/stdc++.h>
using namespace std;

class Solution {
  public:
    void dfs(int idx,vector<vector<int>> &adj,vector<bool> &visited,stack<int> &st){
        visited[idx] = true;
        for(int next: adj[idx]){
            if(!visited[next]) {
                dfs(next,adj,visited,st);
            }
        }
        st.push(idx);
    }
    vector<int> topoSort(int V, vector<vector<int>>& edges) {
        vector<vector<int>> adj(V,vector<int>());
        
        for(auto edge: edges) 
            adj[edge[0]].push_back(edge[1]);
        
        stack<int> st;
        vector<bool> visited(V,false);
        
        for(int i=0;i<V;i++){
            if(!visited[i]){
                dfs(i,adj,visited,st);
            }
        }
        vector<int> res;
        while(!st.empty()){
            res.push_back(st.top());
            st.pop();
        }
        return res;
    }
};

int main(){
    Solution sol;
    vector<vector<int>> edges = {
        {2, 3}, {3, 1}, {4, 0}, {4, 1}, {5, 0}, {5, 2}
    };
    vector<int> result = sol.topoSort(6, edges);
    for(int node: result){
        cout<<node<<" ";
    }
    cout<<endl;
    return 0;
}