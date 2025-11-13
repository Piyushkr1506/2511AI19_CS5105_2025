/*
Find Eventual Safe States - BFS - Topological Sort: G-25 
Problem Statement: A directed graph of V vertices and E edges is given in the form of an 
adjacency list adj. Each node of the graph is labeled with a distinct integer in the range 0 to V - 
1. A node is a terminal node if there are no outgoing edges. A node is a safe node if every 
possible path starting from that node leads to a terminal node. You have to return an array 
containing all the safe nodes of the graph. The answer should be sorted in ascending order. 
Time Complexity: O(V + E + V log V)
Space Complexity: O(V + E) 
*/

#include <bits/stdc++.h>
using namespace std;

class Solution {
  public:
    vector<int> eventualSafeNodes(int V, vector<int> adj[]) {
        vector<int> res;
        vector<int> indegree(V,0);
        vector<vector<int>> revAdj(V,vector<int>());
        queue<int> q;
        
        for(int i=0;i<V;i++){
            for(int node: adj[i]){
                revAdj[node].push_back(i);   
            }
            indegree[i] = adj[i].size();
            if(indegree[i] == 0) {
                q.push(i);
            }
        }
        
        while(!q.empty()){
            int front = q.front();q.pop();
            res.push_back(front);
            for(int neigh: revAdj[front]){
                indegree[neigh] -= 1;
                if(indegree[neigh] == 0) {
                    q.push(neigh);
                }
            }
        }
        sort(res.begin(),res.end());
        return res;
    }
};

int main(){
    int V = 7, E = 7;
    vector<vector<int>> edges = {
        {0, 1},
        {0, 2},
        {1, 2},
        {1, 3},
        {3, 0},
        {4, 5},
        {2, 5}
    };
    vector<int> adj[V];
    for(int i = 0; i < E; i++){
        adj[edges[i][0]].push_back(edges[i][1]);
    }
    Solution obj;
    vector<int> safeNodes = obj.eventualSafeNodes(V, adj);
    for(int node : safeNodes){
        cout << node << " ";
    }
    cout << endl;
    return 0;
}