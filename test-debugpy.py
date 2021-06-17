import debugpy

debugpy.listen(5678)
debugpy.wait_for_client()

breakpoint()

for i in range(0, 100):
    print(i)

1 / 0
