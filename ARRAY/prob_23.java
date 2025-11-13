package Arrays;

import java.util.Scanner;
/*
Kadane's Algorithm

Problem Statements: Given an integer array nums, find the subarray with the largest sum and return the sum of the elements present in that subarray.

A subarray is a contiguous non-empty sequence of elements within an array.
 */
public class prob_23 {

    static int KadaneAlgo(int[] arr ,int n){
        int sum=0;
        int maxSum=Integer.MIN_VALUE;
        for(int i=0;i<n;i++){
            sum+=arr[i];
            if(sum>maxSum){
                maxSum=sum;
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
        int sol = KadaneAlgo(arr,N);
        System.out.println("maximum subarray sum is : "+sol);
        sc.close();

    }
    
}
