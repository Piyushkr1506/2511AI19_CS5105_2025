package Arrays;

import java.util.Arrays;
import java.util.Scanner;

/*Sort an Array of 0’s 1’s and 2’s

Problem Statements: Given an array nums consisting of only 0, 1, or 2. Sort the array in non-decreasing order. The sorting must be done in-place, without making a copy of the original array.


Example 1:
Input: nums = [1, 0, 2, 1, 0]

Output: [0, 0, 1, 1, 2]

Explanation: The nums array in sorted order has 2 zeroes, 2 ones and 1 two
 */
public class prob_22 {
    static void sortArrayZeroOneTwo(int[] arr,int n){
        int zero=0;
        int one=0;
        int two=0;
        for(int i=0;i<n;i++){
            if(arr[i]==0) zero++;
            if(arr[i]==1) one++;
            if(arr[i]==2) two++;
        }
        int j=0;
        while(zero>0){
            arr[j]=0;
            j++;
            zero--;
        }
        while(one>0){
            arr[j]=1;
            j++;
            one--;
        }
        while(two>0){
            arr[j]=2;
            j++;
            two--;
        }
        System.out.println("Array after sort\n"+Arrays.toString(arr));
    }
    public static void main(String[] args) {
         Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }
        System.out.println("initial arr\n"+Arrays.toString(arr));
        sortArrayZeroOneTwo(arr,N);
        sc.close();
    }
}
