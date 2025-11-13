package Arrays;


import java.util.Arrays;
import java.util.HashSet;
import java.util.Scanner;

/*
3 Sum
Problem Statements: Given an integer array nums. Return all triplets such that:
i != j, i != k, and j != k
nums[i] + nums[j] + nums[k] == 0.

Notice that the solution set must not contain duplicate triplets. One element can be a part of multiple triplets. 
The output and the triplets can be returned in any order.

Example 1:
Input: nums = [2, -2, 0, 3, -3, 5]
Output: [[-2, 0, 2], [-3, -2, 5], [-3, 0, 3]]
Explanation: nums[1] + nums[2] + nums[0] = 0
nums[4] + nums[1] + nums[5] = 0
nums[4] + nums[2] + nums[3] = 0
 */

public class prob_30 {
static class Triplets {
    int i;
    int j;
    int k;

    Triplets(int i, int j, int k) {
        int[] arr={i,j,k};
        Arrays.sort(arr);
        this.i = arr[0];
        this.j = arr[1];
        this.k = arr[2];
    }
}
    static void bruteForce3Sum(int arr[],int n){
        HashSet<Triplets> hs = new HashSet<>();
        for(int i=0;i<n;i++){
            for(int j=i+1;j<n;j++){
                for(int k=j+1;k<n;k++){
                    if(i != j&& i != k && j != k && arr[i]+arr[j]+arr[k]==0){
                        // System.out.println(i+" "+j+" "+k);
                        hs.add(new Triplets(arr[i],arr[j],arr[k]));
                    }
                }
            }
        }
        System.out.println();
    }
    public static void main(String[] args) {

         Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }

        bruteForce3Sum(arr,N);
        
    }
}
