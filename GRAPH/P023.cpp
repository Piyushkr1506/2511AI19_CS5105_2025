/*Kahn's Algorithm: Topological Sort
Problem Statement: Given a Directed Acyclic Graph (DAG) with V vertices and E edges, Find 
any Topological Sorting of that Graph. 
Note: In topological sorting, node u will always appear before node v if there is a directed edge 
from node u towards node v(u -> v).
Time Complexity: O(V + E)
*/

#include<bits/stdc++.h>
using namespace std;

class Solution {
  public:
    vector<int> topologicalSort(vector<vector<int>>& adj) {
        int size = adj.size();
        vector<int> res;
        vector<int> indegree(size,0);
        for(const auto list: adj){
            for(int idx: list){
                indegree[idx] += 1;
            }
        }
        queue<int> q;
        for(int i=0;i<size;i++) if(indegree[i]==0) q.push(i);
        while(!q.empty()){
            int top = q.front(); q.pop();
            res.push_back(top);
            for(int it: adj[top]){
                indegree[it]--;
                if(indegree[it]==0) q.push(it);
            }
        }
        return res;
    }
};

int main(){
    Solution sol;
    vector<vector<int>> edges = {
        {2, 3}, {3, 1}, {4, 0}, {4, 1}, {5, 0}, {5, 2}
    };
    vector<vector<int>> adj(6);
    for(auto edge: edges){
        adj[edge[0]].push_back(edge[1]);
    }
    vector<int> result = sol.topologicalSort(adj);
    for(int node: result){
        cout<<node<<" ";
    }
    cout<<endl;
    return 0;
}