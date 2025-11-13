package Arrays;

import java.util.Scanner;

/*
Longest subarray with given sum k (positives)

Problem Statements: Given an array nums of size n and an integer k, find the length of the longest sub-array that sums to k. If no such sub-array exists, return 0.


Example 1:
Input: nums = [10, 5, 2, 7, 1, 9],  k=15

Output: 4

Explanation:

The longest sub-array with a sum equal to 15 is [5, 2, 7, 1], which has a length of 4. This sub-array starts at index 1 and ends at index 4, and the sum of its elements (5 + 2 + 7 + 1) equals 15. Therefore, the length of this sub-array is 4.
 */
public class prob_19 {

    static void sumArraySumBetterForce(int arr[],int N,int targetSum){
        int longestLength=0;
        for(int i=0;i<N;i++){
            int sum=0;
            for(int j=i;j<N;j++){
                    sum+=arr[j];  
           if(sum==targetSum){
                    longestLength=Math.max(longestLength,j-i+1);
                }
            }
            
        }
        System.out.println(longestLength);
    }
    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }
        System.out.println("enter the targetsum");
        int targetSum = sc.nextInt();
        sumArraySumBetterForce(arr,N,targetSum);
        sc.close();
    }
    
}
