/*
Number of Islands - II - Online Queries 
Problem Statement: You are given an n, m which means the row and column of the 2D matrix, 
and an array of size k denoting the number of operations. Matrix elements are 0 if there is water 
or 1 if there is land. Originally, the 2D matrix is all 0 which means there is no land in the matrix. 
The array has k operator(s) and each operator has two integers A[i][0], A[i][1] means that you 
can change the cell matrix[A[i][0]][A[i][1]] from sea to island. Return how many islands are there 
in the matrix after each operation. You need to return an array of size k. 
Note: An island means a group of 1s such that they share a common side. 
Pre-requisite: Disjoint Set data structure 
Time Complexity: O(k * α(n*m)) where α is the Inverse Ackermann function.
*/

class DSU{
    vector<int> parent;
    public:
    DSU(int size){
        parent.resize(size);
        for(int i=0;i<size;i++) {
            parent[i] = i;
        }
    }
    int find(int node){
        if(parent[node] == node) return node;
        return parent[node] = find(parent[node]);
    }
    void union_(int node1, int node2){
        int p1 = find(node1), p2 = find(node2);
        if(p1 != p2) {
            parent[p2] = p1;
        }
    }
};

class Solution {
  public:
    vector<pair<int,int>> dirs = { {-1,0}, {0,-1}, {1,0}, {0,1}};
    vector<int> numOfIslands(int n, int m, vector<vector<int>> &operators) {
        vector<vector<bool>> visited(n,vector<bool>(m,false));
        DSU dsu(n*m);
        vector<int> res;
        int count = 0;
        
        for(auto ops: operators){
            int row = ops[0], col = ops[1];
            int idx1 = row*m + col;
            if(visited[row][col]) {
                res.push_back(count);
                continue;
            }
            visited[row][col] = true;
            count += 1;
            for(auto &[x,y]: dirs) {
                int nx = row+x, ny = col+y;
                int idx2 = nx*m + ny;
                if(nx>=0 and nx<n and ny>=0 and ny<m and visited[nx][ny]) {
                    if(dsu.find(idx1) != dsu.find(idx2)) {
                        count -= 1;
                        dsu.union_(idx1,idx2);
                    }
                }
            }
            res.push_back(count);
        }
        return res;
    }
};

int main(){
    Solution sol;
    vector<vector<int>> operators = {{0,0}, {0,1}, {1,0}, {1,1}, {2,2}};
    vector<int> result = sol.numOfIslands(3, 3, operators);
    for(int count : result) {
        cout << count << " ";
    }
    return 0;
}