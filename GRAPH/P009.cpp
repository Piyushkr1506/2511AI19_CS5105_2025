/*
Number of Ways to Arrive at Destination 
You are in a city that consists of n intersections numbered from 0 to n - 1 with bi-directional 
roads between some intersections. The inputs are generated such that you can reach any 
intersection from any other intersection and that there is at most one road between any two 
intersections. 
You are given an integer n and a 2D integer array ‘roads’ where roads[i] = [ui, vi, timei] means 
that there is a road between intersections ui and vi that takes timei minutes to travel. You want 
to know in how many ways you can travel from intersection 0 to intersection n - 1 in the shortest 
amount of time. 
Return the number of ways you can arrive at your destination in the shortest amount of time. 
Since the answer may be large, return it modulo 109 + 7.
*/

#include<bits/stdc++.h>
using namespace std;

#define ll long long
#define ppii pair<int,pair<int,int>>
#define pii pair<int,int>
const int mod = 1e9+7;

class Solution {
public:
    int countPaths(int n, vector<vector<int>>& roads) {
        vector<ll> dist(n,LLONG_MAX);
        vector<int> count(n,0);
        vector<vector<pii>> adj(n);
        for(auto road: roads){
            int i=road[0],j=road[1],wt=road[2];
            adj[i].push_back({j,wt});
            adj[j].push_back({i,wt});
        }
        priority_queue<pii,vector<pii>,greater<pii>> pq;
        pq.push({0,0});
        dist[0] = 0;
        count[0] = 1;
        while(!pq.empty()){
            auto front = pq.top(); pq.pop();
            ll node = front.second, wt = front.first;
            for(auto &[next,wt]: adj[node]){
                if(wt + dist[node] < dist[next]){
                    dist[next] = wt + dist[node];
                    count[next] = count[node];
                    pq.push({dist[next],next});
                }else if(wt+dist[node] == dist[next]){
                    count[next] = (count[node] + count[next])%mod;
                }
            }
        }
        return count[n-1]%mod;
    }
};

int main(){
    Solution sol;
    int n = 7;
    vector<vector<int>> roads = {
        {0,6,7},{0,1,2},{1,2,3},{1,3,3},{6,3,3},{3,5,1},{6,5,1},{2,5,1},{0,4,5},{4,6,2}
    };
    int result = sol.countPaths(n, roads);
    cout << "Number of ways to arrive at destination: " << result << endl;
    return 0;
}