/**
Number of Provinces
Problem Statement: Given an undirected graph with V vertices. We say two vertices u and v belong to a single province if there is a path from u to v or v to u. Your task is to find the number
of provinces.
Pre-req: Connected Components, Graph traversal techniques
Time Complexity: O(V^2)
Approach: DFS, Disjoint Set
*/

#include<bits/stdc++.h>
using namespace std;

class Solution {
  public:
    void dfs(vector<vector<int>> &adj, vector<bool> & visited, int idx,int V){
        visited[idx] = true;
        for(int i = 0;i<V;i++){
            if(adj[idx][i] && !visited[i]){
                dfs(adj,visited,i,V);
                visited[i] = true;
            }
        }
    }
    int numProvinces(vector<vector<int>> adj, int V) {
        vector<bool> visited(V,false);
        int count = 0;
        for(int i=0;i<V;i++){
            if(!visited[i]){
                dfs(adj,visited,i,V);
                count += 1;
            }
        }
        return count;
    }
};

int main(){
    Solution sol;
    vector<vector<int>> adj = { {1, 0, 1},{0, 1, 0},  {1, 0, 1}};
    int V = 3;
    cout << sol.numProvinces(adj, V) << endl;
    return 0;
}