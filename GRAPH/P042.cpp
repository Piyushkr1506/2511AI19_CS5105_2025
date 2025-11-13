/*Articulation Point in Graph 
Problem Statement: Given an undirected connected graph with V vertices and adjacency list 
adj. You are required to find all the vertices removing which (and edges through it) disconnect 
the graph into 2 or more components. 
Note: Indexing is zero-based i.e nodes numbering from (0 to V-1). There might be loops present 
in the graph. 
Pre-requisite: Bridges in Graph problem & DFS algorithm.
Time Complexity: O(V + E) where V is the number of vertices and E is the number of edges in the graph.
*/

class Solution {
  public:
    void dfs(int curr,vector<int> adj[], vector<int> &disc,vector<int> &low,vector<int> &parent,vector<bool> &articulation){
        static int timer = 0;
        disc[curr] = low[curr] = timer;
        timer += 1;
        int child = 0;
        for(int next: adj[curr]) {
            if(disc[next] == -1) {
                parent[next] = curr;
                child += 1;
                dfs(next,adj,disc,low,parent,articulation);
                low[curr] = min(low[curr],low[next]);
                if(parent[curr]==-1 && child>1) {
                    articulation[curr] = true;
                }
                else if(parent[curr] != -1 && low[next]>=disc[curr]) {
                    articulation[curr] = true;
                }
            }else if(next != parent[curr]) {
                low[curr] = min(low[curr],disc[next]);
            }
        }
    }
    vector<int> articulationPoints(int V, vector<int> adj[]) {
        vector<int> disc(V,-1), low(V,-1),  parent(V,-1), res;
        vector<bool> articulation(V,false);
        
        for(int i=0;i<V;i++){
            if(disc[i] == -1) 
                dfs(i,adj,disc,low,parent,articulation);
        }
        
        for(int i=0;i<V;i++)
            if(articulation[i]) 
                res.push_back(i);
        if(res.size()==0) return {-1};
        return res;
    }
};