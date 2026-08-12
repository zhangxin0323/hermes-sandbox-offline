1. 资源管理器进入补丁文件夹，空白处右键-在终端中打开，粘贴如下命令并回车，开始构建
docker build -t hermes-sandbox-python312:20260812v1 -f Dockerfile.patch .
2. 构建完成后，修改hermes-data\config.yaml的【terminal:】部分，
原始：docker_image: hermes-sandbox-python312:20260810v2
替换为：docker_image: hermes-sandbox-python312:20260812v1
3. 重启hermes，如果是使用docker compose，资源管理器进入hermes目录，资源管理器进入补丁文件夹，空白处右键-在终端中打开
先运行 docker compose down
再运行 docker compose up -d
3. 删除旧沙箱（因为是在旧沙箱的基础上构建的补丁，一定要确认新沙箱构建完成可用了，再删除）
docker rmi docker_image: hermes-sandbox-python312:20260810v2