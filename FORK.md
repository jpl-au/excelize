This repository is a fork of [Excelize](https://github.com/qax-os/excelize). The module path is `github.com/jpl-au/excelize/v2`.

Excelize is the work of its authors and is used under the BSD 3-Clause licence in `LICENSE`. The licence and the copyright notices stay in this fork. The additions are offered under the same licence, and Excelize is welcome to take any of them.

This fork makes no stability promise.

## Updating this fork from Excelize

`go fix` was run on this fork, so expect merge conflicts.

1. Merge the latest Excelize code.

   ```
   git fetch upstream
   git merge upstream/master
   ```

2. For each file with a conflict, keep the Excelize copy.

   ```
   git checkout --theirs <file>
   ```

3. Reapply our own changes to that file, if any.

   ```
   git diff HEAD -- <file>
   ```

4. Run the usual tidy sequence until `go fix` changes nothing more. All tests must pass before the update is committed.

   ```
   go fmt ./... && goimports -w $(git ls-files --cached --others --exclude-standard '*.go') && go fix ./... && go vet ./... && go build ./... && go test ./...
   ```
