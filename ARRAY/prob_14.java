package Arrays;

import java.util.Scanner;
/* 
Longest Consecutive Sequence in an Array.
Problem Statements: Given an array nums of n integers.
Return the length of the longest sequence of consecutive integers. The integers in this sequence can appear in any order.

Example1 :
Input: nums = [100, 4, 200, 1, 3, 2]
Output: 4
Explanation:
The longest sequence of consecutive elements in the array is [1, 2, 3, 4], which has a length of 4. This sequence can be formed regardless of the initial order of the elements in the array.

Example 2:
Input: nums = [0, 3, 7, 2, 5, 8, 4, 6, 0, 1]
Output: 9
Explanation:
The longest sequence of consecutive elements in the array is [0, 1, 2, 3, 4, 5, 6, 7, 8], which has a length of 9.
*/

public class prob_14 {

    static boolean LS(int[] arr,int target){
        for(int i=0;i<arr.length;i++){
            if(target==arr[i]){
                return true;
            }
        }
        return false;
    }
    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }

        int largeSeq=0;
        int count=0;
        for(int i=0;i<N;i++){
            int ele =arr[i];
            count=1;
            while(LS(arr,ele+1)==true){
                ele=ele+1;
                count++;
            }
        }
        if(count>largeSeq){
            largeSeq=count;
            System.out.println(largeSeq);
        }
        sc.close();

    }
}
