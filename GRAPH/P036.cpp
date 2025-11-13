/*
Prim's Algorithm - Minimum Spanning Tree 
Problem Statement: Given a weighted, undirected, and connected graph of V vertices and E 
edges. The task is to find the sum of weights of the edges of the Minimum Spanning Tree. 
(Sometimes it may be asked to find the MST as well, where in the MST the edge-informations 
will be stored in the form {u, v}(u = starting node, v = ending node).
Time Complexity: O(E+ElogV) - Using Priority Queue
Space Complexity: O(V^2) - For the adjacency matrix
*/

#include<bits/stdc++.h>
using namespace std;

#define pii pair<int,int>

class Solution {
  public:
    int spanningTree(int V, vector<vector<int>>& edges) {
        priority_queue<pii,vector<pii>,greater<pii>> pq;
        vector<vector<int>> graph(V,vector<int>(V,1e9));
        vector<bool> visited(V,false);
        
        for(auto edge: edges){
            int st = edge[0], end = edge[1] , wt = edge[2];
            graph[st][end] = wt;
            graph[end][st] = wt;
        }
        
        int sum = 0;
        pq.push({0,0});
        
        while(!pq.empty()){
            auto [curr_dist,curr] = pq.top(); pq.pop();
            if(visited[curr]) continue;
            visited[curr] = true;
            sum += curr_dist;
            for(int i=0;i<V;i++){
                if(graph[curr][i] != 1e9 && !visited[i]) {
                    pq.push({graph[curr][i],i});
                }
            }
        }
        return sum;
    }
};

int main(){
    int V = 5, E = 7;
    vector<vector<int>> edges = {
        {0,1,2},
        {0,3,6},
        {1,2,3},
        {1,3,8},
        {1,4,5},
        {2,4,7},
        {3,4,9}
    };
    
    Solution obj;
    cout<<"Sum of weights of edges of the Minimum Spanning Tree is: "<<obj.spanningTree(V,edges)<<endl;
    return 0;
}