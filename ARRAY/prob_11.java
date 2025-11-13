package Arrays;

import java.util.Scanner;

/*
Kadane’s Algorithm

Problem Statements: Given an integer array nums, find the subarray with the largest sum and return the sum of the elements present in that subarray.

A subarray is a contiguous non-empty sequence of elements within an array.

Example 1:
Input: nums = [2, 3, 5, -2, 7, -4]
Output: 15
Explanation: The subarray from index 0 to index 4 has the largest sum = 15

Example 2:
Input: nums = [-2, -3, -7, -2, -10, -4]
Output: -2
Explanation: The element on index 0 or index 3 make up the largest sum when taken as a subarray.

 */
public class prob_11 {

    static int KadanesAlgorithm(int nums[],int n){
        int sum=0;
        int maxSum=Integer.MIN_VALUE;
        // int startI =-1;
        // int endI =-1;
        // int start=-1;
        for(int i=0;i<n;i++){
            
            // if(sum==0) {
            //     start=i;// used for finding the start and end of subarray
            // }
            sum+=nums[i];
            if(sum>maxSum){
                maxSum=sum;
                // startI=start;// used for finding the start and end of subarray
                // endI=i;
            }
            if(sum<0){
                sum=0;
            }
        }
        return maxSum;
    }
   
    public static void main(String[] args) {
         Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }

        int maxSum = KadanesAlgorithm(arr,N);
        System.out.println("the largest sum  of the elements present in that subarray:"+maxSum);
        sc.close();
    }
    
}
