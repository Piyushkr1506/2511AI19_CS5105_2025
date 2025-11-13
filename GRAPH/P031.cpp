/*
Cheapest Flights Within K Stops 
There are n cities and m edges connected by some number of flights. You are given an array of 
flights where flights[i] = [ fromi, toi, pricei] indicates that there is a flight from city fromi to city toi 
with cost price. You have also given three integers src, dst, and k, and return the cheapest price 
from src to dst with at most k stops. If there is no such route, return -1.
Time Complexity: O(n) n: number of flights
*/
#include<bits/stdc++.h>
using namespace std;

class Solution {
  public:
    const int INF = 1e9+5;
    int CheapestFLight(int n, vector<vector<int>>& flights, int src, int dst, int k) {
        vector<int> dist(n,INF);
        queue<pair<int,pair<int,int>>> q;
        vector<vector<pair<int,int>>> adj(n,vector<pair<int,int>>());
        
        for(auto flight: flights) {
            adj[flight[0]].push_back({flight[1],flight[2]});
        }
        
        dist[src] = 0;
        q.push({0,{src,0}});
        
        while(!q.empty()){
            auto front = q.front(); q.pop();
            int stops = front.first, u = front.second.first, wt = front.second.second;
            if(stops>k) continue;
            for(auto &[v,cost]: adj[u]){
                if(wt+cost < dist[v] and stops<=k){
                    dist[v] = wt + cost;
                    q.push({stops+1,{v,dist[v]}});
                }
            }
        }
        return dist[dst]==INF?-1:dist[dst];
    }
};

int main(){
    Solution sol;
    int n = 4, m = 5 ;
    vector<vector<int>> flights ={{0,1,100},{1,2,100},{2,0,100},{1,3,600},{2,3,200}} ;
    int src=0, dst=3, k=1 ;
    cout<<sol.CheapestFLight(n,flights,src,dst,k);
    return 0;
}