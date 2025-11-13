package Arrays;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Scanner;

/*
Leaders in an array.

Problem Statements: Given an integer array nums, return a list of all the leaders in the array.

A leader in an array is an element whose value is strictly greater than all elements to its right in the given array. The rightmost element is always a leader. 
The elements in the leader array must appear in the order they appear in the nums array.

Example 1:
Input: nums = [1, 2, 5, 3, 1, 2]
Output: [5, 3, 2]
Explanation: 2 is the rightmost element, 3 is the largest element in the index range [3, 5], 5 is the largest element in the index range [2, 5]

Example 2:
Input: nums = [-3, 4, 5, 1, -4, -5]
Output: [5, 1, -4, -5]
 */
public class prob_13 {

    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }
        System.out.println(Arrays.toString(arr));
        ArrayList<Integer> al =new ArrayList<>();
        //Brute force
        // for(int i=0;i<N;i++){
        //     boolean leader = true;
        //     for(int j=i+1;j<N;j++){
        //         if(arr[j]>arr[i]){
        //             leader=false;
        //         }
        //     }
        //     if(leader){
        //         al.add(arr[i]);
        //     }
        // }
        // System.out.println(al);

        //optimal solution

        int maxArray = Integer.MIN_VALUE;
        for(int i=N-1;i>=0;i--){
            if(arr[i]>maxArray){
                al.add(arr[i]);
            }
            maxArray = Math.max(arr[i],maxArray);
        }
        Collections.reverse(al);
        System.out.println(al);
        sc.close();
    }
    
}
