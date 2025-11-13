/*
Bridges in Graph - Using Tarjan's Algorithm of time in and low time 
Problem Statement: There are n servers numbered from 0 to n - 1 connected by undirected 
server-to-server connections forming a network where connections[i] = [ai, bi] represents a 
connection between servers ai and bi. Any server can reach other servers directly or indirectly 
through the network. 
A critical connection is a connection that, if removed, will make some servers unable to reach 
some other servers. 
Return all critical connections in the network in any order. 
Note: Here servers mean the nodes of the graph. The problem statement is taken from 
leetcode. 
Pre-requisite: DFS algorithm 
Time Complexity: O(V + E) where V is the number of vertices and E is the number of edges in the graph.
*/

class Solution{
    void dfs(int curr, vector<vector<int>> &adj,vector<int> &disc,vector<int> &low,
    vector<int> &parent,  vector<vector<int>> &res){
        static int timer = 0;
        disc[curr] = low[curr] = timer;
        timer += 1;
        for(int next: adj[curr]){
            if(disc[next] == -1){
                parent[next] = curr;
                dfs(next,adj,disc,low,parent,res);
                low[curr] = min(low[curr],low[next]);
                if(low[next] > disc[curr]) {
                    res.push_back({curr,next});
                }
            }
            else if(next != parent[curr]) {
                low[curr] = min(low[curr],disc[next]);
            }
        }
    }

    vector<vector<int>> findBridges(vector<vector<int>> &edges, int V, int e) {
        vector<vector<int>> res,adj(V,vector<int>());
        vector<int> disc(V,-1), low(V,-1),  parent(V,-1);

        for(auto edge: edges){
            int st = edge[0], end = edge[1];
            adj[st].push_back(end);
            adj[end].push_back(st);
        }

        for(int i=0;i<V;i++) {
            if(disc[i] == -1) 
                dfs(i,adj,disc,low,parent,res);
        }

        return res;
    }
};