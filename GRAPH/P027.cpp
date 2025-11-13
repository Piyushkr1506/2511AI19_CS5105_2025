/*
Alien Dictionary - Topological Sort 
Problem Statement: Given a sorted dictionary of an alien language having N words and k 
starting alphabets of a standard dictionary. Find the order of characters in the alien language. 
Note: Many orders may be possible for a particular test case, thus you may return any valid 
order.
Time Complexity: O(N*K) where N is the number of words and K is the number of starting alphabets.
*/

#include<bits/stdc++.h>
using namespace std;

class Solution {
public:
    string findOrder(vector<string> &words) {
        vector<vector<int>> adj(26);
        vector<int> indegree(26, 0);
        vector<bool> present(26, false);
        
        // Mark all characters as present
        for (auto &w : words) {
            for (char c : w) present[c - 'a'] = true;
        }

        // Build graph using first differing character
        for (int i = 0; i + 1 < words.size(); i++) {
            string &a = words[i], &b = words[i + 1];
            int len = min(a.size(), b.size());
            bool found = false;

            for (int j = 0; j < len; j++) {
                if (a[j] != b[j]) {
                    int u = a[j] - 'a', v = b[j] - 'a';
                    adj[u].push_back(v);
                    found = true;
                    break;
                }
            }

            // Invalid case: prefix issue (e.g. "abc" before "ab")
            if (!found && a.size() > b.size()) return "";
        }

        // Compute indegree
        for (int u = 0; u < 26; u++) {
            for (int v : adj[u]) indegree[v]++;
        }

        // Topological sort (Kahn’s Algorithm)
        queue<int> q;
        for (int i = 0; i < 26; i++) {
            if (present[i] && indegree[i] == 0) q.push(i);
        }

        string order = "";
        while (!q.empty()) {
            int u = q.front(); q.pop();
            order += (char)('a' + u);
            for (int v : adj[u]) {
                if (--indegree[v] == 0) q.push(v);
            }
        }

        // Check if all present chars were used
        int total_chars = count(present.begin(), present.end(), true);
        if (order.size() != total_chars) return ""; // cycle or incomplete

        return order;
    }
};

int main(){
    Solution sol;
    vector<string> words = {"baa", "abcd", "abca","baab", "caba", "cad"};
    string order = sol.findOrder(words);
    if (order.empty()) {
        cout << "No valid order exists." << endl;
    } else {
        cout << "The order of characters is: " << order << endl;
    }
}