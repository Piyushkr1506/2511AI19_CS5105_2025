package Arrays;

import java.util.Scanner;

/*Find Missing Number.

Problem Statements: Given an integer array of size n containing distinct values in the range from 0 to n (inclusive), 
return the only number missing from the array within this range.

Example1:
Input: nums = [0, 2, 3, 1, 4]

Output: 5
 */
public class prob_8 {
    
    static int FindMissingNumber(int[] nums,int n){
        int sum = (n*(n+1))/2;
        int numSum=0;
        for(int i=0;i<n;i++){
            numSum+=nums[i];
        }
        return sum-numSum;
    }
    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }
        int ans = FindMissingNumber(arr,N);
        System.out.println("The Missing number of the given range is: "+ans);
        sc.close();
    }
}
