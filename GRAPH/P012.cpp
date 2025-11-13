/*
Depth First Search (DFS)
Problem Statement: Given an undirected graph, return a vector of all nodes by traversing the
graph using depth-first search (DFS). Complete Traversal in case of a disconnected graph.
Pre-req: Recursion, Graph Representation
Time Complexity: O(V + E) where V is the number of vertices and E is the number of edges
Space Complexity: O(V) for the visited array and O(V) for the recursion stack in worst case
*/

#include<bits/stdc++.h>
using namespace std;

class Solution {
  public:
    void dfs(vector<vector<int>>& adj,int idx, vector<int> &res, vector<int> &visited){
        visited[idx] = true;
        res.push_back(idx);
        for(int next: adj[idx]){
            if(!visited[next]){
                dfs(adj,next,res,visited);
                visited[next] = true;
            }
        }
    }
    vector<int> dfs(vector<vector<int>>& adj) {
        vector<int> res;
        int v = adj.size();
        vector<int> visited(v,false);
        for(int i=0;i<v;i++){
            if(!visited[i]) {
                dfs(adj,i,res,visited);
            }
        }
        return res;
    }
};

int main(){
    Solution sol;
    vector<vector<int>> adj = {
        {2, 3, 1}, {0}, {0, 4}, {0}, {2}
    };
    vector<int> res = sol.dfs(adj);
    for(int it: res) cout<<it<<" ";
    return 0;
}