/*
Strongly Connected Components - Kosaraju's Algorithm 
Problem Statement: Given a Directed Graph with V vertices (Numbered from 0 to V-1) and E 
edges, Find the number of strongly connected components in the graph. 
Pre-requisite: DFS algorithm
Time complexity: O(V + E)
*/
class Solution {
  public:
    void dfs(int curr,vector<vector<int>> &adj,vector<bool> &visited,stack<int> &st){
        visited[curr] = true;
        for(const int node: adj[curr]){
            if(!visited[node]){
                dfs(node,adj,visited,st);
            }
        }
        st.push(curr);
    }
    void dfs(int curr,vector<vector<int>> &adj,vector<bool>&visited,vector<int> &nodes){
        visited[curr] = true;
        for(const int node: adj[curr]){
            if(!visited[node]){
                dfs(node,adj,visited,nodes);
            }
        }
        nodes.push_back(curr);
    }
    int kosaraju(vector<vector<int>> &adj) {
        int size = adj.size();
        /*storing order of traversal based on ending time*/
        vector<bool> visited(size,false);
        stack<int> st;
        for(int i=0;i<size;i++){
            if(visited[i]==false){
                dfs(i,adj,visited,st);
            }
        }
        /*reversing the graph*/
        vector<vector<int>> newadj(size,vector<int>());
        for(int i=0;i<size;i++){
            for(const int it: adj[i]){
                newadj[it].push_back(i);
            }
        }
        /*traversing to find strongly connected components*/
        fill(visited.begin(),visited.end(),false);
        vector<vector<int>> result;
        while(!st.empty()){
            int curr = st.top(); st.pop();
            if(!visited[curr]){
                vector<int> temp;
                dfs(curr,newadj,visited,temp);
                result.push_back(temp);
            }
        }
        return result.size();
    }
};
