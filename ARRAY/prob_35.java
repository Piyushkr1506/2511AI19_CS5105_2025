package Arrays;

import java.util.Scanner;

/*Count subarrays with given xor K

Problem Statements: Given an array of integers nums and an integer k, return the total number of subarrays whose XOR equals to k.

Example 1:
Input : nums = [4, 2, 2, 6, 4], k = 6
Output : 4
Explanation : The subarrays having XOR of their elements as 6 are [4, 2],  [4, 2, 2, 6, 4], [2, 2, 6], and [6]

Example 2:
Input :nums = [5, 6, 7, 8, 9], k = 5
Output : 2
Explanation : The subarrays having XOR of their elements as 5 are [5] and [5, 6, 7, 8, 9
 */
public class prob_35 {

    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        System.out.println("enter the target value k:");
        int k=sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }
        // XorEqTarget(arr,N,k);

        sc.close();
    }
}
