/**
 * BFS (Breadth-First Search) Implementation
Problem Statement: Given an undirected graph, return a vector of all nodes by traversing the graph using breadth-first search (BFS).
Pre-req: Graph Representation, Queue STL
Time Complexity: O(V + E)
Approach: Iterative BFS
 */

#include<bits/stdc++.h>
using namespace std;

class Solution {
  public:
    vector<int> bfs(vector<vector<int>> &adj) {
        int size = adj.size();
        vector<bool> visited(size,false);
        vector<int> res;
        queue<int> q; q.push(0);
        while(!q.empty()){
            int curr = q.front(); q.pop();
            res.push_back(curr);
            visited[curr] = true;
            for(int next: adj[curr]){
                if(!visited[next]){
                    q.push(next);
                    visited[next] = true;
                }
            }
        }
        return res;
    }
};

int main() {
    int vertex, edges;
    cin >> vertex >> edges;
    vector<vector<int>> adj(vertex + 1);
    for (int i = 0; i < edges; i++) {
        int u, v;
        cin >> u >> v;
        adj[u].push_back(v);
        adj[v].push_back(u);
    }
    
    Solution sol = Solution();
    vector<int> bfs = sol.bfs(adj);

    cout<<"BFS Traversal: ";
    for(int i : bfs) cout << i << " ";
    cout << endl;

    return 0;
}