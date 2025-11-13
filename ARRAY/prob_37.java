package Arrays;

import java.util.Scanner;

/*Count Inversions

Problem Statements: Given an integer array nums. Return the number of inversions in the array.
Two elements a[i] and a[j] form an inversion if a[i] > a[j] and i < j.
1.It indicates how close an array is to being sorted.
2.A sorted array has an inversion count of 0.
3.An array sorted in descending order has maximum inversion.
Example 1:
Input: nums = [2, 3, 7, 1, 3, 5]
Output: 5
Explanation: The responsible indexes are:
nums[0], nums[3], values: 2 > 1 & indexes: 0 < 3
nums[1], nums[3], values: 3 > 1 & indexes: 1 < 3
nums[2], nums[3], values: 7 > 1 & indexes: 2 < 3
nums[2], nums[4], values: 7 > 3 & indexes: 2 < 4
nums[2], nums[5], values: 7 > 5 & indexes: 2 < 5
 */
public class prob_37 {

    static int InversionCount(int[] arr,int n){
        int count=0;
        for(int i=0;i<n;i++){
            for(int j=i+1;j<n;j++){
                if(i<j && arr[i]>arr[j]){
                    count++;
                }
            }
        }
        return count;
    }
    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }
        int ans = InversionCount(arr,N);
        System.out.println("inversion count:"+ans);
        sc.close();
    }
    
}
