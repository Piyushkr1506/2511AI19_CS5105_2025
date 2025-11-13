/*
Accounts Merge 
Problem Statement: Given a list of accounts where each element account [ i ] is a list of strings, 
where the first element account [ i ][ 0 ]  is a name, and the rest of the elements are emails 
representing emails of the account. 
Now, we would like to merge these accounts. Two accounts definitely belong to the same 
person if there is some common email to both accounts. Note that even if two accounts have 
the same name, they may belong to different people as people could have the same name. A 
person can have any number of accounts initially, but all of their accounts definitely have the 
same name. 
After merging the accounts, return the accounts in the following format: the first element of each 
account is the name, and the rest of the elements are emails in sorted order. 
Note: Accounts themselves can be returned in any order. 
Pre-requisite: Disjoint Set data structure 
*/

// User function Template for C++

class DSU {
public:
    vector<int> parent, rankv;
    DSU(int size) {
        parent.resize(size);
        rankv.resize(size, 0);
        for (int i = 0; i < size; i++) parent[i] = i;
    }
    int find(int node) {
        if (parent[node] == node) return node;
        return parent[node] = find(parent[node]); // path compression
    }
    void union_(int u, int v) {
        int pu = find(u), pv = find(v);
        if (pu == pv) return;
        if (rankv[pu] < rankv[pv]) parent[pu] = pv;
        else if (rankv[pv] < rankv[pu]) parent[pv] = pu;
        else {
            parent[pv] = pu;
            rankv[pu]++;
        }
    }
};


class Solution {
public:
    vector<vector<string>> accountsMerge(vector<vector<string>> &accounts) {
        unordered_map<string, int> emailToIndex;
        unordered_map<string, string> emailToName;
        int idx = 0;

        // Step 1: Map each unique email to an index and store name
        for (auto &acc : accounts) {
            string name = acc[0];
            for (int i = 1; i < acc.size(); i++) {
                string email = acc[i];
                if (emailToIndex.find(email) == emailToIndex.end()) {
                    emailToIndex[email] = idx++;
                    emailToName[email] = name;
                }
            }
        }

        // Step 2: Create DSU of all unique emails
        DSU dsu(idx);

        // Step 3: Union all emails in same account
        for (auto &acc : accounts) {
            string firstEmail = acc[1];
            int firstIdx = emailToIndex[firstEmail];
            for (int i = 2; i < acc.size(); i++) {
                dsu.union_(firstIdx, emailToIndex[acc[i]]);
            }
        }

        // Step 4: Group emails by parent
        unordered_map<int, vector<string>> parentEmails;
        for (auto &[email, i] : emailToIndex) {
            int parent = dsu.find(i);
            parentEmails[parent].push_back(email);
        }

        // Step 5: Sort and build final result
        vector<vector<string>> res;
        for (auto &[par, emails] : parentEmails) {
            sort(emails.begin(), emails.end());
            string name = emailToName[emails[0]];
            vector<string> temp;
            temp.push_back(name);
            temp.insert(temp.end(), emails.begin(), emails.end());
            res.push_back(temp);
        }

        return res;
    }
};


int main(){
    Solution sol;
    vector<vector<string>> accounts = {
        {"John", "johnsmith@mail.com", "john00@mail.com"},
        {"John", "johnnybravo@mail.com"},
        {"Mary", "mary@mail.com"},
        {"John", "john_newyork@mail.com"}
    };
    vector<vector<string>> mergedAccounts = sol.accountsMerge(accounts);
    for (auto &acc : mergedAccounts) {
        cout << "[";
        for (auto &s : acc) {
            cout << s << ", ";
        }
        cout << "]" << endl;
    }
    return 0;
}