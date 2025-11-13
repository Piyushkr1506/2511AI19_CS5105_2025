package Arrays;

import java.util.Scanner;

/*
Two Sum 

Problem Statements: Given an array of integers nums and an integer target. Return the indices(0 - indexed) of two elements in nums such that they add up to target.

Each input will have exactly one solution, and the same element cannot be used twice. Return the answer in increasing order.


Example1:
Input: nums = [1, 6, 2, 10, 3], target = 7

Output: [0, 1]
 */
public class prob_21 {
    static void twoSum(int[] arr,int N,int target){
            for(int i=0;i<N;i++){
                for(int j=1;j<N;j++){
                    if(arr[i]+arr[j]==target){
                        System.out.println("the index are :"+"{"+i+","+j+"}");
                    }
                }
            }    
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
        int target =sc.nextInt();
        twoSum(arr,N,target);
        sc.close();
    }
}
