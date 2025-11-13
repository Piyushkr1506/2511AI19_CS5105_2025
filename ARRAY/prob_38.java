package Arrays;

import java.util.Scanner;

/*
Reverse Pairs
Problem Statements: Given an integer array nums. Return the number of reverse pairs in the array.
An index pair (i, j) is called a reverse pair if:

1. 0 <= i < j < nums.length
2. nums[i] > 2 * nums[j]

Example 1:
Input: nums = [6, 4, 1, 2, 7]
Output: 3

Explanation: The reverse pairs are:
(0, 2) : nums[0] = 6, nums[2] = 1, 6 > 2 * 1
(0, 3) : nums[0] = 6, nums[3] = 2, 6 > 2 * 2
(1, 2) : nums[1] = 4, nums[2] = 1, 4 > 2 * 1


Example 2:
Input: nums = [5, 4, 4, 3, 3]
Output: 0
Explanation: No pairs satisfy both the conditions.
 */
public class prob_38 {

    static int BrutereversePair(int[] arr,int n){
        
        int countRev=0;
        for(int i=0;i<n;i++){
            for(int j=i+1;j<n;j++){
                if(i < j &&  arr[i] > 2 * arr[j]){
                countRev=countRev+1;
                }
            }
        }
         return countRev;   
        
    }
static void merge(int arr[],int l,int m,int h){
        int temp[] = new int[h];
        int left = l;
        int right= m+1;
        while(l<=h && left<=m && right<=h){

            if(arr[left]<=arr[right]){
                temp[l]=arr[left];
                left++;
                l++;
            }
            else{
                temp[l]=arr[right];
                right++;
                l++;
            }
            
        }
        while(l<h && left<=m){
            temp[l]=arr[left];
            left++;
            l++;
        }
        while(l<h  && right<=h){
            temp[l]=arr[right];
            right++;
            l++;
        }

    }
    static int reversePair(int arr[],int l,int m,int h){
        int count=0;
        int right=m+1;
        for(int i=l;i<=m;i++){
            while(right <= h && arr[i] > 2*arr[right]){
                right++;
                count+=(right-(m+1));
            }
        }
        return count;
    }
    static int mergesort(int arr[],int l,int h){
        int count=0;
        int mid=h+l/2;
        if(l>=h) return count;
        count+=mergesort(arr, l, mid);
        count+=mergesort(arr, mid+1, h);
        count+=reversePair(arr,l,mid,h);
        merge(arr,l,mid,h);
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

        
        // int answer = reversePair(arr,N);
        System.out.println("Number of reverse pair in given array: "+mergesort(arr,0,N));
        sc.close();
    }
}
