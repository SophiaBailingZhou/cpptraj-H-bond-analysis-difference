This Excel macro takes in `cpptraj` data from two different simulations (e.g. wildtype-ligand binding and mutant ligand binding) and creates a color-coded list for the change of bond strength for each H-bond.  The changes are given through time-weighted average distance and time-weighted average angle.

# Requirements
- MS Excel
- cpptraj data of two different simulations

# Detailed explanation of what the script does
The macro consists of two scripts that need to be executed. `Sort_Hbond_Analysis_part1` and `Sort_Hbond_Analysis_part2`.

Sort_Hbond_Analysis_part1 does the following:
1.	Create column that concatenates the H bond donor and acceptor columns, this serves as a unique identifier of each H bond
2.	Remove duplicate rows that exist in both the acceptor and donor mask outputs
3.	Sort by concatenated column

Sort_Hbond_Analysis_part2 does the following:
1.	Align WT and variant entries via the unique concatenated H bond identifier
2.	Sort both tables by concatenated column
⦁	this is important for H bond summing in step 3 – any chemically identical H bonds (e.g. the different H bonds a Glu side chain can form with its two carboxy oxygens) should be in adjacent rows after this
3.	(optional user-aided summing of H bonds. The script parses the donor and acceptor columns and suggests summing up a row with the next row if the next donor/acceptor value is identical. The second row is deleted then, and a merge log counter is increased by one. Additionally, the deleted concatenated identifier is saved in a log as well. The script will also always automatically create a sheet without the summed H bonds.)
4.	Calculates differences between WT and variant (frames, frequency, frequency-weighted average distance, frequency-weighted average angle)
5.	Format all data as excel-table with sortable columns
6.	Add color-scale to ΔFrames for easier visualisation of large changes

# Acknowledgements
The utilised xlAlignkeys submodule was created by excel forum user shg.
