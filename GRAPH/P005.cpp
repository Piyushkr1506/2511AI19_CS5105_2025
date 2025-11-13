/*
Bipartite Graph | DFS Implementation
Problem Statement: Given an adjacency list of a graph adj of V no. of vertices having 0 based
index. Check whether the graph is bipartite or not.
If we are able to colour a graph with two colours such that no adjacent nodes have the same
colour, it is called a bipartite graph.
*/

#include <bits/stdc++.h>
using namespace std;

class Solution {
  public:
    bool isBipartite(int V, vector<vector<int>> &edges) {
        vector<int> color(V,-1);
        vector<vector<int>> adj(V);
        for(auto edge: edges){
            int i = edge[0], j = edge[1];
            adj[i].push_back(j);
            adj[j].push_back(i);
        }
        queue<int> q;
        q.push(0); 
        color[0] = 0; // marking vertex-0 as color-0
        while(!q.empty()){
            int node = q.front(); q.pop();
            for(int next: adj[node]){
                if(color[next]==-1){
                    color[next] = !color[node];
                    q.push(next);
                }else if(color[next] == color[node]) { // if adjacent nodes have same color return false;
                    return false;
                }
            }
        }
        return true;
    }
};

int main(){
    Solution sol;
    int V = 5;
    vector<vector<int>> edges = {{0,1},{0,3},{1,2},{1,4},{2,3}};
    if(sol.isBipartite(V, edges)){
        cout << "Graph is Bipartite" << endl;
    }else{
        cout << "Graph is not Bipartite" << endl;
    }
    return 0;
}