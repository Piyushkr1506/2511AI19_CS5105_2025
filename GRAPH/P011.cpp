/*
Most Stones Removed with Same Row or Column - DSU: G-53
Problem Statement: There are n stones at some integer coordinate points on a 2D plane. Each
coordinate point may have at most one stone.
You need to remove some stones.
A stone can be removed if it shares either the same row or the same column as another stone
that has not been removed.
Given an array of stones of length n where stones[i] = [xi, yi] represents the location of the ith
stone, return the maximum possible number of stones that you can remove.
Pre-requisite: Disjoint Set data structure
Time Complexity: O(N * α(N)) where N is the number of stones and α is the Inverse Ackermann function
Space Complexity: O(N)
*/

#include<bits/stdc++.h>
using namespace std;

class DisjointSet{
    public:
    vector<int> parent;
    vector<int> size_;
    DisjointSet() {}
    DisjointSet(int size){
        parent.resize(size);
        size_.resize(size);
        for(int i=0;i<size;i++){
            parent[i] = i; size_[i] = 1;
        }
    }
    int find(int idx){
        if(idx == parent[idx]) return idx;
        return parent[idx] = find(parent[idx]);
    }
    void union_by_size(int p1, int p2){
        if(size_[p1]>size_[p2]){
            parent[p2] = p1;
            size_[p1] += size_[p2];
        }else{
            parent[p1] = p2;
            size_[p2] += size_[p1];
        }
    }
};

class Solution {
public:
    int removeStones(vector<vector<int>>& stones) {
        int max_row = INT_MIN, max_col = INT_MIN;
        int stone_count = stones.size();
        for(auto stone: stones) {
            max_row = max(max_row,stone[0]);
            max_col = max(max_col,stone[1]);
        }
        int total_size = max_row + max_col + 2;
        DisjointSet ds(total_size);
        for(auto stone: stones){
            int p1 = ds.find(stone[0]), p2 = ds.find(stone[1]+max_row+1);
            if(p1 != p2) {
                ds.union_by_size(p1,p2);
            }
        }
        int components = 0;
        for(int i=0;i<total_size;i++){
            if(ds.parent[i] == i and ds.size_[i]>1) components += 1;
        }
        return stone_count - components;
    }
};

int main(){
    Solution sol;
    vector<vector<int>> stones = {{0,0},{0,1},{1,0},{1,2},{2,1},{2,2}};
    cout<<sol.removeStones(stones);
    return 0;
}