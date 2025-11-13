/*
Making a Large Island 
Problem Statement: You are given an n x n binary grid. A grid is said to be binary if every value 
in the grid is either 1 or 0. You can change at most one cell in the grid from 0 to 1. You need to 
find the largest group of connected  1's. Two cells are said to be connected if both are adjacent 
to each other and both have the same value. 
Pre-requisite: Disjoint Set data structure 
Time complexity: O(n*n*α(n*n)) where α is the Inverse Ackermann function.
*/

class DSU{
    vector<int> parent;
    public:
    vector<int> size_;
    DSU(int size){
        parent.resize(size);
        size_.resize(size);
        for(int i=0;i<size;i++) {
            parent[i] = i;
            size_[i] = 1;
        }
    }
    int find(int node){
        if(parent[node] == node) return node;
        return parent[node] = find(parent[node]);
    }
    void union_(int u, int v){
        int pu = find(u), pv = find(v);
        if (pu == pv) return;
        if (size_[pu] < size_[pv]) swap(pu, pv);
        parent[pv] = pu;
        size_[pu] += size_[pv];
    }
};

class Solution {
    vector<pair<int,int>> dirs = { {-1,0}, {0,-1}, {1,0}, {0,1}};
  public:
    int largestIsland(vector<vector<int>>& grid) {
        int n = grid.size();
        DSU dsu(n*n);
        
        for(int i=0;i<n;i++){
            for(int j=0;j<n;j++){
                if(grid[i][j] == 0) continue;
                int idx1 = i*n + j;
                for(auto &[x,y]: dirs) {
                    int nx = i+x, ny = j+y;
                    int idx2 = nx*n + ny;
                    if(nx>=0 and nx<n and ny>=0 and ny<n and grid[nx][ny]) {
                        if(dsu.find(idx1) != dsu.find(idx2)) {
                            dsu.union_(idx1,idx2);
                        }
                    }
                }
            }
        }
        int max_ = INT_MIN;
        for(int i=0;i<n;i++){
            for(int j=0;j<n;j++){
                if(grid[i][j] == 0) {
                    int sum = 0;
                    set<int> st;
                    for(auto &[x,y]: dirs) {
                        int nx = i+x, ny = j+y;
                        int idx2 = nx*n + ny;
                        if(nx>=0 and nx<n and ny>=0 and ny<n and grid[nx][ny]) {
                            st.insert(dsu.find(idx2));
                        }
                    }
                    for(int it: st) sum += dsu.size_[it];
                    max_ = max(max_,sum);
                }
            }
        }
        return max_==INT_MIN?n*n:max_+1;
    }
};

int main(){
    Solution sol;
    vector<vector<int>> grid = {
        {1, 0, 1},
        {0, 1, 0},
        {1, 0, 1}
    };
    int result = sol.largestIsland(grid);
    cout << "Largest Island Size: " << result << endl;
    return 0;
}