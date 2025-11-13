/*
There are a total of numCourses courses you have to take, labeled from 0 to numCourses - 1. You are given an array prerequisites where prerequisites[i] = [ai, bi] indicates that you must take course bi first if you want to take course ai.
For example, the pair [0, 1], indicates that to take course 0 you have to first take course 1.
Return the ordering of courses you should take to finish all courses. If there are many valid answers, return any of them. If it is impossible to finish all courses, return an empty array.
Time Complexity: O(V + E)
Space Complexity: O(V) for queue and indegree array
*/
#include<bits/stdc++.h>
using namespace std;

class Solution {
public:
    vector<int> findOrder(int n, vector<vector<int>>& preq) {
        vector<vector<int>> adj(n);
        vector<int> indegree(n,0);
        for(const auto edge: preq){
            adj[edge[1]].push_back(edge[0]);
            indegree[edge[0]] += 1;
        }
        int count = 0;
        vector<int> order;
        queue<int> q; 
        for(int i=0;i<indegree.size();i++){
            if(indegree[i]==0) q.push(i);
        }
        while(!q.empty()){
            int top = q.front(); q.pop();
            order.push_back(top);
            count += 1;
            for(int it: adj[top]){
                indegree[it] -= 1;
                if(indegree[it]==0) q.push(it);
            }
        }
        return count == n?order:vector<int>();
    }
};

int main(){
    Solution sol;
    vector<vector<int>> preq = {
        {1,0},{2,0},{3,1},{3,2}
    };
    vector<int> result = sol.findOrder(4, preq);
    if(result.empty()){
        cout << "Cannot finish all courses" << endl;
    } else {
        cout << "Course order: ";
        for(int course: result){
            cout << course << " ";
        }
        cout << endl;
    }
    return 0;
}