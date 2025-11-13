/*
Shortest Path in Directed Acyclic Graph Topological Sort
Given a DAG, find the shortest path from the source to all other nodes in this DAG. In this 
problem statement, we have assumed the source vertex to be ‘0’. You will be given the weighted 
edges of the graph. 
Note: What is a DAG ( Directed Acyclic Graph)? 
A Directed Graph (containing one-sided edges) having no cycles is said to be a Directed Acyclic 
Graph. 
Time Complexity: 
O(VE)- Bellman-Ford Algorithm
O(E+VlogV) - dijkstra's algorithm
O(V+E) - Using Topological Sort
*/

#include<bits/stdc++.h>
using namespace std;

class Solution {
    public:
    const int INF = 1e9+5;
    vector<int> shortestPath(int V, int E, vector<vector<int>>& edges) {
        vector<int> dist(V,INF);
        vector<vector<pair<int,int>>> adj(V,vector<pair<int,int>>());
        queue<int> q;
        vector<int> indegree(V,0);
        vector<int> order;
        
        for(auto edge: edges){
            int u = edge[0], v = edge[1], w = edge[2];
            adj[u].push_back({v,w});
            indegree[v] += 1;
        }
        
        dist[0] = 0;
        for(int i=0;i<V;i++){
            if(indegree[i] == 0) q.push(i);
        }
        while(!q.empty()){
            int curr = q.front(); q.pop();
            order.push_back(curr);
            for(auto &[next,weight]: adj[curr]){
                indegree[next] -= 1;
                if(indegree[next] == 0) {
                    q.push(next);
                }
            }
        }
        
        for(int curr: order){
            for(auto &[next,wt]: adj[curr]) {
                if(dist[curr]+wt < dist[next]) {
                    dist[next] = dist[curr] + wt;
                }
            }
        }
        
        for(int &it: dist) 
            if(it==INF) 
                it = -1;
        
        return dist;
    }
};


int main(){
    Solution sol;
    int n = 6, m = 7 ;
    vector<vector<int>> edges ={{0,1,2},{0,4,1},{1,2,3},{4,2,2},{4,5,4},{2,3,6},{5,3,1}} ;
    vector<int> dist = sol.shortestPath(n,m,edges);
    for(auto it: dist) cout<<it<<" ";
    return 0;
}