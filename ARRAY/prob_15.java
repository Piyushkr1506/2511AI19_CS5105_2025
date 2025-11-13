package Arrays;

import java.util.Scanner;

/*Count subarrays with given sum.

Problem Statements: Given an array of integers nums and an integer k, return the total number of subarrays whose sum equals to k.


Example 1:
Input: nums = [1, 1, 1], k = 2

Output: 2

Explanation: In the given array [1, 1, 1], there are two subarrays that sum up to 2: [1, 1] and [1, 1]. Hence, the output is 2.

Example 2:
Input: nums = [1, 2, 3], k = 3

Output: 2

Explanation: In the given array [1, 2, 3], there are two subarrays that sum up to 3: [1, 2] and [3]. Hence, the output is 2.
 */
public class prob_15 {
    static void sumArraySumBruteForce(int arr[],int N,int targetSum){
        int count=0;
        for(int i=0;i<N;i++){
            for(int j=i;j<N;j++){
                int sum=0;
                for(int k=i;k<=j;k++){
                    sum+=arr[k];
                }
                if(sum==targetSum){
                    count++;
                }
            }
        }
        System.out.println(count);
    }
    static void sumArraySumBetterForce(int arr[],int N,int targetSum){
        int count=0;
        for(int i=0;i<N;i++){
            int sum=0;
            for(int j=i;j<N;j++){
                    sum+=arr[j];  
           if(sum==targetSum){
                    count++;
                }
            }
            
        }
        System.out.println(count);
    }

    static void sumArraySumOptimal(int[] arr,int N,int k){
        
    }
    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }
        
        System.out.println("enter the sum required for subarray");
        int targetSum =sc.nextInt();
        
        //brute force
        // sumArraySumBruteForce(arr,N,targetSum);
        //better
        // sumArraySumBetterForce(arr,N,targetSum);
        //optimal
        sumArraySumOptimal(arr,N,targetSum);
        sc.close();
    }
}
