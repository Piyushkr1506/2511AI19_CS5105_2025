package Arrays;


import java.util.Arrays;
import java.util.Scanner;

/*Rearrange array elements by sign

Problem Statements: Given an integer array nums of even length consisting of an equal number of positive and negative integers.Return the answer array in such a way that the given conditions are met:


1.Every consecutive pair of integers have opposite signs.
2.For all integers with the same sign, the order in which they were present in nums is preserved.
3.The rearranged array begins with a positive integer.

Example 1:
Input : nums = [2, 4, 5, -1, -3, -4]

Output : [2, -1, 4, -3, 5, -4]

Explanation: The positive number 2, 4, 5 maintain their relative positions and -1, -3, -4 maintain their relative positions
 */
public class prob_24 {


    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }
        // int ans[] = OppositeSign(arr,N);
        // System.out.println(Arrays.toString(ans));
        sc.close();
    }
}
