/*
There are a total of numCourses courses you have to take, labeled from 0 to numCourses - 1. You are given an array prerequisites where prerequisites[i] = [ai, bi] indicates that you must take course bi first if you want to take course ai.
For example, the pair [0, 1], indicates that to take course 0 you have to first take course 1.
Return true if you can finish all courses. Otherwise, return false.
Time Complexity: O(V + E)
Space Complexity: O(V) for queue and indegree array
*/
#include<bits/stdc++.h>
using namespace std;

class Solution {
public:
    bool canFinish(int n, vector<vector<int>>& preq) {
        vector<vector<int>> adj(n);
        vector<int> indegree(n,0);
        for(const auto edge: preq){
            adj[edge[0]].push_back(edge[1]);
            indegree[edge[1]] += 1;
        }
        int count = 0;
        queue<int> q; 
        for(int i=0;i<indegree.size();i++){
            if(indegree[i]==0) q.push(i);
        }
        while(!q.empty()){
            int top = q.front(); q.pop();
            count += 1;
            for(int it: adj[top]){
                indegree[it] -= 1;
                if(indegree[it]==0) q.push(it);
            }
        }
        return count == n;
    }   
};

int main(){
    Solution sol;
    vector<vector<int>> preq = {
        {1,0},{2,0},{3,1},{3,2}
    };
    bool result = sol.canFinish(4, preq);
    cout << (result ? "Can finish all courses" : "Cannot finish all courses") << endl;
    return 0;
}