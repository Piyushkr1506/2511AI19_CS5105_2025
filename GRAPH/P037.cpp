/*
Kruskal's Algorithm - Minimum Spanning Tree 
Problem Statement: Given a weighted, undirected, and connected graph of V vertices and E 
edges. The task is to find the sum of weights of the edges of the Minimum Spanning Tree. 
Time Complexity: O(ElogE) - Using Sorting and Union-Find
Space Complexity: O(V) - For the parent and rank arrays
*/
#include<bits/stdc++.h>
using namespace std;

class Solution {
  public:
    int find(int idx, vector<int> &parent){
        if(parent[idx] == -1) return idx;
        return parent[idx] = find(parent[idx],parent);
    }
    void union_(int a,int b, vector<int> &parent){
        parent[b] = a;
    }
    bool detect_cycle(int a,int b, vector<int> &parent){
        int s1 = find(a,parent);
        int s2 = find(b,parent);
        if(s1 != s2) {
            union_(s1,s2,parent);
            return false;
        }
        return true;
    }
    int kruskalsMST(int V, vector<vector<int>> &edges) {
        sort(edges.begin(),edges.end(),[](auto &a, auto &b){
            return a[2]<b[2];
        });
        vector<int> parent(V,-1);
        int mst = 0;
        for(auto edge: edges){
            int u = edge[0], v = edge[1], w = edge[2];
            if(!detect_cycle(u,v,parent)){
                mst += w;
            }
        }
        return mst;
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
    cout<<"Sum of weights of edges of the Minimum Spanning Tree is: "<<obj.kruskalsMST(V,edges)<<endl;
    return 0;
}