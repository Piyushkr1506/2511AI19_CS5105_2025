/*
Word Ladder - I  
Given are the two distinct words startWord and targetWord, and a list denoting wordList of 
unique words of equal lengths. Find the length of the shortest transformation sequence from 
startWord to targetWord. 
In this problem statement, we need to keep the following conditions in mind: 
A word can only consist of lowercase characters. 
Only one letter can be changed in each transformation. 
Each transformed word must exist in the wordList including the targetWord. 
startWord may or may not be part of the wordList 
Note:  If there’s no possible way to transform the sequence from startWord to targetWord return 
0.
Pre-req: Graphs, BFS on Graphs
Time Complexity: O(N*M*M)
*/

#include<bits/stdc++.h>
using namespace std;

class Solution {
  public:
    int wordLadderLength(string start, string target, vector<string>& words) {
        int length = 1;
        set<string> st;
        for(auto word: words) st.insert(word);
        queue<string> q;
        q.push(start);
        st.erase(start);
        while(!q.empty()){
            int size = q.size();
            for(int k=0;k<size;k++){
                string curr = q.front(); q.pop();
                if(curr == target) return length;
                for(int i=0;i<curr.length();i++){
                    char original = curr[i];
                    for(char ch = 'a';ch<='z';ch++){
                        curr[i] = ch;
                        if(st.find(curr) != st.end()){
                            q.push(curr);
                            st.erase(curr);
                        }
                    }
                    curr[i] = original;
                }
            }
            length += 1;
        }
        return 0;
    }
};

int main(){
    string startWord = "hit", targetWord = "cog";
    vector<string> wordList = {"hot","dot","dog","lot","log","cog"};
    Solution obj;
    int ans = obj.wordLadderLength(startWord,targetWord,wordList);
    cout<<ans<<endl;
    return 0;
}