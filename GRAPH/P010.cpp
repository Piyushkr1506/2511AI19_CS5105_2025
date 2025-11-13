/*
Problem Statement: You are given a graph with n vertices and m edges. You can remove one
edge from anywhere and add that edge between any two vertices in one operation. Find the
minimum number of operations that will be required to make the graph connected. If it is not
possible to make the graph connected, return -1.
Pre-requisite: Disjoint Set data structure
Time Complexity: O(E + V) using DFS, O(N) using DSU
*/

#include<bits/stdc++.h>
using namespace std;

class Solution {
public:
    int find(vector<int>&parent, int idx){
        if(parent[idx] == -1) {
            return idx;
        }
        return parent[idx] = find(parent,parent[idx]);
    }
    void union_(vector<int>&parent, int p1, int p2){
        parent[p1] = p2;
    }
    int makeConnectedUsingDSU(int n, vector<vector<int>>& cons) {
        vector<int> parent(n,-1);
        int extra = 0;
        for(auto con: cons){
            int p1 = find(parent,con[0]), p2 = find(parent,con[1]);
            if(p1 == p2){
                extra += 1;
            }else{
                union_(parent,p1,p2);
            }
        }
        int req = 0;
        for(int it: parent) { if(it==-1) req += 1; }
        if(extra<req-1) return -1;
        return req-1;
    }
    void dfs(vector<vector<int>> &adj, int idx, vector<bool> &visited){
        visited[idx] = true;
        for(int next: adj[idx]){
            if(!visited[next]){
                dfs(adj,next,visited);
                visited[next] = true;
            }
        }
    }
    int minimumConnectionsUsingDFS(int n, vector<vector<int>> &cons) {
        int components = 0;
        vector<bool> visited(n,false);
        int total_edge = cons.size();
        if(total_edge < n-1) return -1;
        vector<vector<int>> adj(n);
        for(auto con: cons){
            int u = con[0], v = con[1];
            adj[u].push_back(v);
            adj[v].push_back(u);
        }
        for(int i=0;i<n;i++){
            if(!visited[i]){
                components += 1;
                dfs(adj,i,visited);
            }
        }
        int extra = total_edge - (n - components);
        if(extra<components-1) return -1;
        return components-1;
    }
};

int main(){
    Solution sol;
    int n = 6;
    vector<vector<int>> connections = {
        {0,1},{0,2},{0,3},{1,2}, {1,3}
    };
    int result = sol.makeConnectedUsingDSU(n, connections);
    cout << "Minimum operations to make graph connected (using DSU): " << result << endl;
    result = sol.minimumConnectionsUsingDFS(n, connections);
    cout << "Minimum operations to make graph connected (using DFS): " << result << endl;
    return 0;
}