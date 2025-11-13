package Arrays;

import java.util.Scanner;

/*Maximum Consecutive Ones.

Problem Statements: Given a binary array nums, return the maximum number of consecutive 1s in the array.

A binary array is an array that contains only 0s and 1s.

Example1:
Input: nums = [1, 1, 0, 0, 1, 1, 1, 0]

Output: 3

Explanation: The maximum consecutive 1s are present from index 4 to index 6, amounting to 3 1s

Example 2:
Input: nums = [0, 0, 0, 0, 0, 0, 0, 0]

Output: 0

Explanation: No 1s are present in nums, thus we return 0
 */
public class prob_9 {

    static int MaxConsecutiveOnes(int[] num,int n){
        int ones=0;
        int maxOnes=0;
        for(int i=0;i<n;i++){
            if(num[i]==1){
                ones++;
                maxOnes= Math.max(maxOnes,ones);
            }
            else{
                ones=0;
            }
        }
        return maxOnes;
    }
    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }
        int sol =MaxConsecutiveOnes(arr,N);
        System.out.println("Maximum number of consecutive ones in given array are : "+sol);
        sc.close();
    }
    
}
